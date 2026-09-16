#!/usr/bin/env python
"""
ppt2video: PPT/PDF 슬라이드 + 페이지별 대본 -> TTS 음성 -> 영상(mp4) 자동 생성기

사용 흐름
---------
1. PDF 파일에서 각 페이지를 이미지로 렌더링한다 (슬라이드 비주얼).
2. 페이지별 대본을 얻는다:
   - --pptx 로 원본 PPTX를 주면 슬라이드 노트(발표자 노트)를 자동 추출
   - 또는 --script 로 별도 텍스트/JSON 파일을 지정
3. 대본을 오픈소스 TTS(Coqui XTTS v2)로 음성 합성한다.
4. 이미지 + 음성으로 페이지별 클립을 만들고, 순서대로 이어붙여 최종 영상을 만든다.

필요 프로그램: ffmpeg (PATH에 있어야 함)
"""

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path


def eprint(*args, **kwargs):
    print(*args, file=sys.stderr, **kwargs)


# --------------------------------------------------------------------------
# 1. PDF -> 이미지
# --------------------------------------------------------------------------

def render_pdf_pages(pdf_path: Path, out_dir: Path, target_width: int = 1920):
    try:
        import pymupdf as fitz  # PyMuPDF (new import name)
    except ImportError:
        import fitz  # PyMuPDF (legacy import name)

    out_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    image_paths = []
    for i, page in enumerate(doc):
        rect = page.rect
        zoom = target_width / rect.width
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        img_path = out_dir / f"page_{i+1:03d}.png"
        pix.save(str(img_path))
        image_paths.append(img_path)
    doc.close()
    return image_paths


# --------------------------------------------------------------------------
# 2. 대본 확보
# --------------------------------------------------------------------------

def extract_pptx_notes(pptx_path: Path):
    from pptx import Presentation

    prs = Presentation(str(pptx_path))
    notes = []
    for slide in prs.slides:
        if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
            text = slide.notes_slide.notes_text_frame.text.strip()
        else:
            text = ""
        notes.append(text)
    return notes


def load_script_file(script_path: Path):
    if script_path.suffix.lower() == ".json":
        data = json.loads(script_path.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            raise ValueError("--script JSON 파일은 문자열 리스트여야 합니다. 예: [\"1페이지 대본\", \"2페이지 대본\"]")
        return [str(x).strip() for x in data]
    else:
        raw = script_path.read_text(encoding="utf-8")
        # 페이지 구분자는 "===" 로 시작하는 줄, 없으면 빈 줄 두 번 연속으로 구분
        if "===" in raw:
            parts = [p.strip() for p in raw.split("\n===") ]
            parts = [p.lstrip("=").strip() for p in parts]
        else:
            parts = [p.strip() for p in raw.split("\n\n")]
        return [p for p in parts]


def resolve_scripts(args, num_pages: int):
    scripts = None
    if args.script:
        scripts = load_script_file(Path(args.script))
    elif args.pptx:
        scripts = extract_pptx_notes(Path(args.pptx))
    else:
        eprint("오류: --script 또는 --pptx 중 하나는 반드시 지정해야 합니다 (대본 출처).")
        sys.exit(1)

    if len(scripts) != num_pages:
        eprint(
            f"경고: 대본 개수({len(scripts)})와 PDF 페이지 수({num_pages})가 다릅니다. "
            f"짧은 쪽 기준으로 자르고, 모자란 페이지는 대본 없음(무음)으로 처리합니다."
        )
        if len(scripts) < num_pages:
            scripts = scripts + [""] * (num_pages - len(scripts))
        else:
            scripts = scripts[:num_pages]
    return scripts


# --------------------------------------------------------------------------
# 3. TTS
# --------------------------------------------------------------------------

class TtsEngine:
    def __init__(self, language="ko", speaker_wav=None, speaker=None, device=None,
                 speed=1.15, temperature=0.65, repetition_penalty=5.0):
        os.environ.setdefault("COQUI_TOS_AGREED", "1")  # XTTS 라이선스(CPML, 비상업적 이용) 자동 동의
        from TTS.api import TTS
        import torch

        if device is None:
            device = "cuda" if torch.cuda.is_available() else "cpu"
        eprint(f"[TTS] 모델 로딩 중 (device={device})... 최초 실행 시 모델 다운로드로 시간이 걸릴 수 있습니다.")
        self.tts = TTS("tts_models/multilingual/multi-dataset/xtts_v2").to(device)
        self.language = language
        self.speaker_wav = speaker_wav
        self.speaker = speaker
        self.speed = speed
        self.temperature = temperature
        self.repetition_penalty = repetition_penalty

        if not speaker_wav and not speaker:
            available = getattr(self.tts, "speakers", None)
            if available:
                self.speaker = available[0]
                eprint(f"[TTS] --speaker-wav / --speaker 미지정 -> 기본 내장 화자 사용: {self.speaker}")
            else:
                eprint(
                    "[TTS] 경고: 내장 화자를 찾을 수 없고 --speaker-wav 도 없습니다. "
                    "음성 클로닝용 짧은 wav 샘플(6초 이상 권장)을 --speaker-wav 로 지정하세요."
                )

    def synthesize(self, text: str, out_path: Path):
        kwargs = dict(
            text=text,
            language=self.language,
            file_path=str(out_path),
            speed=self.speed,
            temperature=self.temperature,
            repetition_penalty=self.repetition_penalty,
        )
        if self.speaker_wav:
            kwargs["speaker_wav"] = self.speaker_wav
        elif self.speaker:
            kwargs["speaker"] = self.speaker
        self.tts.tts_to_file(**kwargs)

    def list_speakers(self):
        return list(getattr(self.tts, "speakers", []) or [])


def make_silence(out_path: Path, duration: float = 1.2):
    subprocess.run(
        [
            "ffmpeg", "-y", "-f", "lavfi", "-i", f"anullsrc=r=24000:cl=mono",
            "-t", str(duration), str(out_path),
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def get_audio_duration(path: Path) -> float:
    result = subprocess.run(
        [
            "ffprobe", "-v", "error", "-show_entries", "format=duration",
            "-of", "default=noprint_wrappers=1:nokey=1", str(path),
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    return float(result.stdout.strip())


def pad_audio(in_path: Path, out_path: Path, pad_seconds: float):
    subprocess.run(
        [
            "ffmpeg", "-y", "-i", str(in_path),
            "-af", f"apad=pad_dur={pad_seconds}",
            str(out_path),
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


# --------------------------------------------------------------------------
# 4. 영상 합성
# --------------------------------------------------------------------------

def make_page_clip(image_path: Path, audio_path: Path, out_path: Path, fps: int = 25):
    subprocess.run(
        [
            "ffmpeg", "-y",
            "-loop", "1", "-i", str(image_path),
            "-i", str(audio_path),
            "-c:v", "libx264", "-tune", "stillimage",
            "-c:a", "aac", "-b:a", "192k",
            "-pix_fmt", "yuv420p",
            "-r", str(fps),
            "-shortest",
            "-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2",
            str(out_path),
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


def concat_clips(clip_paths, out_path: Path, list_file: Path):
    list_file.write_text(
        "\n".join(f"file '{p.resolve().as_posix()}'" for p in clip_paths),
        encoding="utf-8",
    )
    subprocess.run(
        [
            "ffmpeg", "-y", "-f", "concat", "-safe", "0",
            "-i", str(list_file), "-c", "copy", str(out_path),
        ],
        check=True,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

def build_arg_parser():
    p = argparse.ArgumentParser(description="PDF 슬라이드 + 대본 -> TTS 영상 자동 생성")
    p.add_argument("--pdf", required=False, help="슬라이드 이미지를 뽑아올 PDF 파일 (PowerPoint에서 '내보내기 > PDF'로 생성)")
    p.add_argument("--pptx", required=False, help="발표자 노트를 대본으로 자동 추출할 원본 PPTX 파일 (동일한 덱)")
    p.add_argument("--script", required=False, help="페이지별 대본 파일 (.json 문자열 배열, 또는 '===' 로 페이지 구분한 .txt)")
    p.add_argument("--out", default="output.mp4", help="출력 영상 경로 (기본: output.mp4)")
    p.add_argument("--lang", default="ko", help="TTS 언어 코드 (기본: ko)")
    p.add_argument("--speaker-wav", default=None, help="음성 클로닝용 참조 wav 파일 (6초 이상 권장)")
    p.add_argument("--speaker", default=None, help="XTTS 내장 화자 이름 (--list-speakers 로 확인)")
    p.add_argument("--device", default=None, choices=["cpu", "cuda"], help="TTS 실행 장치 (기본: 자동 감지)")
    p.add_argument("--speed", type=float, default=1.15, help="TTS 발화 속도 배율. 1.0=기본(느리고 늘어짐), 클수록 빠름 (기본 1.15)")
    p.add_argument("--temperature", type=float, default=0.65, help="TTS 생성 다양성/표현력 (기본 0.65, 낮을수록 안정적/단조로움)")
    p.add_argument("--repetition-penalty", type=float, default=5.0, help="같은 소리 반복(질질 끄는 발음) 억제 강도 (기본 5.0)")
    p.add_argument("--pad", type=float, default=0.4, help="각 페이지 음성 뒤 여백(초) (기본 0.4)")
    p.add_argument("--min-duration", type=float, default=1.2, help="대본이 빈 페이지의 노출 시간(초) (기본 1.2)")
    p.add_argument("--width", type=int, default=1920, help="렌더링할 슬라이드 이미지 가로 픽셀 (기본 1920)")
    p.add_argument("--fps", type=int, default=25, help="출력 영상 fps (기본 25)")
    p.add_argument("--workdir", default="build", help="중간 산출물(이미지/오디오/클립) 저장 폴더 (기본: build)")
    p.add_argument("--keep-temp", action="store_true", help="완료 후 중간 산출물을 지우지 않음")
    p.add_argument("--list-speakers", action="store_true", help="XTTS 내장 화자 목록만 출력하고 종료")
    return p


def check_ffmpeg():
    if shutil.which("ffmpeg") is None or shutil.which("ffprobe") is None:
        eprint("오류: ffmpeg / ffprobe 를 찾을 수 없습니다. PATH에 설치되어 있는지 확인하세요.")
        sys.exit(1)


def main():
    args = build_arg_parser().parse_args()
    check_ffmpeg()

    if args.list_speakers:
        engine = TtsEngine(language=args.lang, device=args.device)
        for name in engine.list_speakers():
            print(name)
        return

    if not args.pdf:
        eprint("오류: --pdf 는 필수입니다 (슬라이드 이미지 소스).")
        sys.exit(1)

    workdir = Path(args.workdir)
    img_dir = workdir / "images"
    audio_dir = workdir / "audio"
    clip_dir = workdir / "clips"
    for d in (img_dir, audio_dir, clip_dir):
        d.mkdir(parents=True, exist_ok=True)

    print("[1/4] PDF 페이지를 이미지로 렌더링 중...")
    image_paths = render_pdf_pages(Path(args.pdf), img_dir, target_width=args.width)
    print(f"  -> {len(image_paths)} 페이지")

    print("[2/4] 대본 확보 중...")
    scripts = resolve_scripts(args, len(image_paths))

    print("[3/4] TTS 음성 생성 및 페이지 클립 생성 중...")
    engine = None
    clip_paths = []
    for i, (img_path, text) in enumerate(zip(image_paths, scripts)):
        page_no = i + 1
        raw_wav = audio_dir / f"page_{page_no:03d}_raw.wav"
        final_wav = audio_dir / f"page_{page_no:03d}.wav"
        clip_path = clip_dir / f"clip_{page_no:03d}.mp4"

        if text:
            if engine is None:
                engine = TtsEngine(
                    language=args.lang,
                    speaker_wav=args.speaker_wav,
                    speaker=args.speaker,
                    device=args.device,
                    speed=args.speed,
                    temperature=args.temperature,
                    repetition_penalty=args.repetition_penalty,
                )
            print(f"  - {page_no}/{len(image_paths)} 페이지 TTS 생성 중... ({text[:20]}...)")
            engine.synthesize(text, raw_wav)
            pad_audio(raw_wav, final_wav, args.pad)
        else:
            print(f"  - {page_no}/{len(image_paths)} 페이지: 대본 없음 -> 무음 {args.min_duration}초")
            make_silence(final_wav, args.min_duration)

        make_page_clip(img_path, final_wav, clip_path, fps=args.fps)
        clip_paths.append(clip_path)

    print("[4/4] 전체 영상 합치는 중...")
    concat_clips(clip_paths, Path(args.out), workdir / "concat_list.txt")

    if not args.keep_temp:
        shutil.rmtree(workdir, ignore_errors=True)

    print(f"완료: {args.out}")


if __name__ == "__main__":
    main()
