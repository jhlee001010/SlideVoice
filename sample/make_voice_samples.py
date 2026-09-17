"""모든 내장 목소리를 같은 문장으로 만들어서 sample/voices/ 에 저장 (직접 들어보고 고르기용)."""
from pathlib import Path
from supertonic import TTS

OUT_DIR = Path("sample/voices")
OUT_DIR.mkdir(parents=True, exist_ok=True)

TEXT = "오늘은 트랜스포머 모델의 어텐션 구조에 대해서 차근차근 설명해 드리겠습니다."
SPEED = 1.2  # 강의용으로 조금 빠르게

tts = TTS(auto_download=True)

for voice_name in tts.voice_style_names:
    style = tts.get_voice_style(voice_name=voice_name)
    wav, dur = tts.synthesize(text=TEXT, lang="ko", voice_style=style, total_steps=8, speed=SPEED)
    out_path = OUT_DIR / f"{voice_name}.wav"
    tts.save_audio(wav, str(out_path))
    print(f"{voice_name}: {out_path} ({dur[0]:.2f}s)")
