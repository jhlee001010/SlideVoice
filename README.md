# ppt2video

PPT/PDF 슬라이드 + 페이지별 대본을 오픈소스 TTS(Coqui XTTS v2)로 음성 합성해서
자동으로 나레이션 영상(mp4)을 만들어주는 CLI 도구입니다.

## 동작 방식

1. **슬라이드 이미지**: `--pdf`로 준 PDF의 각 페이지를 이미지로 렌더링합니다.
   (PPTX는 PowerPoint에서 "파일 > 내보내기 > PDF로 만들기"로 미리 변환해주세요.
   LibreOffice 등 별도 변환 프로그램 설치가 필요 없습니다.)
2. **페이지별 대본**: 아래 중 하나로 지정합니다.
   - `--pptx`: 같은 덱의 원본 .pptx를 주면 슬라이드 노트(발표자 노트)를 자동으로 읽어옵니다.
   - `--script`: 별도 파일로 직접 지정합니다.
     - `.json`: 문자열 배열, 예) `["1페이지 대본", "2페이지 대본"]`
     - `.txt`: 페이지 사이를 `===`로 구분 (없으면 빈 줄 두 번으로 구분)
   - PDF 페이지 수와 대본 개수가 다르면 경고를 띄우고 짧은 쪽에 맞춥니다. 대본이 빈 페이지는 무음으로 처리됩니다.
3. **TTS 음성 생성**: Coqui XTTS v2 모델로 각 페이지 대본을 음성으로 합성합니다.
   - 한국어(`--lang ko`) 포함 다국어 지원.
   - 목소리는 `--speaker-wav`(6초 이상의 참조 음성으로 목소리 클로닝) 또는
     `--speaker`(내장 화자 이름, `--list-speakers`로 목록 확인)로 지정. 둘 다 없으면 내장 화자 중 첫 번째를 자동 사용.
4. **영상 합성**: 이미지 + 음성으로 페이지별 클립을 만든 뒤 순서대로 이어붙여 최종 mp4를 만듭니다.

## 설치

```bash
python -m venv venv
venv\Scripts\pip install -r requirements.txt
```

ffmpeg가 시스템 PATH에 있어야 합니다 (Windows: `winget install Gyan.FFmpeg`).

## 사용 예시

```bash
# PPTX 노트를 대본으로 자동 사용
venv\Scripts\python ppt2video.py --pdf deck.pdf --pptx deck.pptx --out output.mp4

# 대본을 별도 파일로 지정 + 목소리 클로닝
venv\Scripts\python ppt2video.py --pdf deck.pdf --script script.json --speaker-wav my_voice.wav --out output.mp4

# 내장 화자 목록 확인
venv\Scripts\python ppt2video.py --list-speakers
```

## 주요 옵션

| 옵션 | 설명 | 기본값 |
|---|---|---|
| `--pdf` | 슬라이드 이미지 소스 PDF (필수) | - |
| `--pptx` | 발표자 노트를 대본으로 자동 추출할 PPTX | - |
| `--script` | 대본 파일 (.json / .txt) | - |
| `--lang` | TTS 언어 코드 | `ko` |
| `--speaker-wav` | 목소리 클로닝용 참조 음성 | - |
| `--speaker` | XTTS 내장 화자 이름 | 자동(첫 번째) |
| `--device` | `cpu` / `cuda` | 자동 감지 |
| `--pad` | 각 페이지 음성 뒤 여백(초) | `0.4` |
| `--min-duration` | 대본 없는 페이지 노출 시간(초) | `1.2` |
| `--width` | 렌더링 슬라이드 가로 해상도 | `1920` |
| `--fps` | 출력 영상 fps | `25` |
| `--workdir` | 중간 산출물 폴더 | `build` |
| `--keep-temp` | 중간 산출물 보존 | 끔 |

## 참고 / 주의사항

- **XTTS v2 모델 라이선스**: Coqui Public Model License(CPML)로, **비상업적 용도**로만 무료 사용 가능합니다.
  상업적으로 쓰려면 별도 라이선스가 필요합니다. 첫 실행 시 모델(~2GB)이 자동 다운로드됩니다.
- CPU만 있어도 동작하지만 GPU(CUDA)가 있으면 훨씬 빠릅니다.
- 목소리 클로닝(`--speaker-wav`) 시 본인 목소리나 사용 권한이 있는 음성만 사용하세요.

## 샘플로 테스트하기

```bash
venv\Scripts\python sample\make_test_pdf.py
venv\Scripts\python ppt2video.py --pdf sample\test.pdf --script sample\script.json --out sample\test_output.mp4
```
