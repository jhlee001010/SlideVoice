# SlideVoice

PPT 또는 PDF 슬라이드와 발표 대본만 있으면, AI TTS 음성을 입힌 MP4 강의 영상을 자동으로 생성합니다.

## 이 프로그램으로 할 수 있는 것

- AI가 자연스러운 음성으로 발표 대본을 읽어주는 나레이션 영상을 만들어줍니다. 직접 마이크로 녹음할 필요가 없습니다.
- 강의 영상, 발표 영상, 온라인 강의 자료 등을 빠르게 제작할 때 사용합니다.
- 핵심 TTS 및 영상 생성 구성요소는 무료 오픈소스 기반이며, 별도의 유료 TTS 서비스 가입이 필요하지 않습니다.

마우스로 클릭하는 프로그램이 아니라, 명령어(글자로 된 지시문)를 한 줄 입력해서 실행하는 방식입니다. 어렵게 느껴질 수 있지만, 아래 설명대로 **복사 → 붙여넣기 → 엔터**만 반복하면 됩니다. 프로그래밍 지식이 없어도 따라 할 수 있게 순서대로 정리해뒀습니다.

Windows / macOS / Linux(WSL 포함) 어디서나 쓸 수 있습니다. 아래에서 내가 쓰는 OS(대부분 Windows) 섹션만 따라가면 됩니다.

---

## 가장 빠르게 사용하는 방법

전체 과정이 한눈에 들어오도록 미리 보여드리면 이렇습니다. (Python/ffmpeg 설치 등 최초 준비 과정은 아래 "Windows" 또는 "macOS/Linux" 섹션을 먼저 따라가세요. 준비가 끝났다면 그다음부터는 이 4단계만 반복하면 됩니다.)

1. PPT 파일을 `SlideVoice` 폴더에 넣습니다.
2. PPT의 발표자 노트에 슬라이드별 대본을 작성합니다. (**반드시 한국어로** — 자세한 요령은 아래 [대본 작성 방법](#대본-작성-방법) 참고)
3. 아래 명령어를 실행합니다.

```powershell
venv\Scripts\python ppt2video.py --pptx deck.pptx --out output.mp4
```

4. 잠시 기다리면 `output.mp4`가 생성됩니다.

---

## Windows

### 1단계. 저장소 받기

```powershell
git clone https://github.com/jhlee001010/SlideVoice.git
cd SlideVoice
```

> git이 없다면 GitHub 페이지에서 "Code > Download ZIP"으로 받아서 압축을 풀어도 됩니다.

### 2단계. Python 설치 확인

```powershell
python --version
```

오류가 나면 설치 (설치 후 PowerShell 창을 새로 열어야 인식됨):

```powershell
winget install -e --id Python.Python.3.12
```

### 3단계. ffmpeg 설치 확인

```powershell
ffmpeg -version
```

오류가 나면 설치 (설치 후 PowerShell 창을 새로 열어야 인식됨):

```powershell
winget install -e --id Gyan.FFmpeg
```

### 4단계. 가상환경 만들고 패키지 설치

(SlideVoice 폴더 안에서 실행)

```powershell
python -m venv venv
venv\Scripts\pip install -r requirements.txt
```

몇 분 정도 걸릴 수 있습니다.

### 5단계. PDF 자동 변환 프로그램 확인

`--pdf` 없이 `.pptx` 파일만 넘기면 PowerPoint를 자동으로 실행해서 PDF로 변환합니다.
**Microsoft 365/Office(PowerPoint)가 설치되어 있으면 별도 설치 없이 바로 됩니다.**

이미 PDF로 변환해둔 파일이 있다면 이 단계는 필요 없습니다 — `--pdf`로 바로 지정하면 이 자동 변환 과정 자체를 건너뛰어서 더 빠르게 실행됩니다 (아래 8단계 예시 참고).

### 6단계. 슬라이드 파일 준비

만든 `.pptx` 파일을 `SlideVoice` 폴더 **바로 안에** 복사해 넣습니다. (예: `deck.pptx`)

슬라이드 노트(발표자 노트)에 각 페이지 대본을 미리 적어두면 자동으로 읽어옵니다. 대본을 어떻게 써야 하는지는 아래 [대본 작성 방법](#대본-작성-방법)을 꼭 먼저 읽어주세요.

### 7단계. 실행!

```powershell
venv\Scripts\python ppt2video.py --pptx deck.pptx --out output.mp4
```

별도 대본 파일(`script.json`)을 쓰는 경우:

```powershell
venv\Scripts\python ppt2video.py --pptx deck.pptx --script script.json --out output.mp4
```

PDF를 이미 직접 만들어뒀다면 `--pdf`로 지정하세요. PPTX → PDF 자동 변환 과정을 건너뛰기 때문에 더 빠르게 실행됩니다:

```powershell
venv\Scripts\python ppt2video.py --pdf deck.pdf --script script.json --out output.mp4
```

실행하는 동안 화면에 `[1/4] ... [2/4] ... [3/4] ... [4/4] ...` 진행 상황이 나오고,
끝나면 같은 폴더에 `output.mp4`가 생성됩니다.

> **파일을 못 찾는다는 오류가 나면**: `--pptx`/`--pdf`로 준 파일이 지금 명령어를 실행하는 폴더 기준 경로에 없다는 뜻입니다.
> 오류 메시지에 찍히는 절대경로를 보고 파일을 그 위치로 옮기거나, 경로를 정확히 적어주세요 (예: `sample\deck.pptx`).

### 8단계 (선택). 먼저 샘플로 테스트해보기

```powershell
venv\Scripts\python sample\make_test_pdf.py
venv\Scripts\python ppt2video.py --pdf sample\test.pdf --script sample\script.json --out sample\test_output.mp4
```

`sample\test_output.mp4`가 생기면 정상 작동하는 것입니다.

---

## macOS / Linux (WSL 포함)

### 1단계. 저장소 받기

```bash
git clone https://github.com/jhlee001010/SlideVoice.git
cd SlideVoice
```

### 2단계. Python 설치 확인

```bash
python3 --version
```

오류가 나면 설치:

```bash
# macOS
brew install python

# Ubuntu/WSL
sudo apt install python3 python3-venv python3-pip
```

### 3단계. ffmpeg 설치 확인

```bash
ffmpeg -version
```

오류가 나면 설치:

```bash
# macOS
brew install ffmpeg

# Ubuntu/WSL
sudo apt install ffmpeg
```

### 4단계. 가상환경 만들고 패키지 설치

(SlideVoice 폴더 안에서 실행)

```bash
python3 -m venv venv
venv/bin/pip install -r requirements.txt
```

몇 분 정도 걸릴 수 있습니다.

### 5단계. PDF 자동 변환 프로그램 설치 (LibreOffice)

`--pdf` 없이 `.pptx` 파일만 넘기면 LibreOffice로 자동 변환합니다.

```bash
# macOS
brew install --cask libreoffice

# Ubuntu/WSL
sudo apt install libreoffice
```

이미 PDF로 변환해둔 파일이 있다면 이 단계는 필요 없습니다 — `--pdf`로 바로 지정하면 이 자동 변환 과정 자체를 건너뛰어서 더 빠르게 실행됩니다 (아래 7단계 예시 참고).

### 6단계. 슬라이드 파일 준비

만든 `.pptx` 파일을 `SlideVoice` 폴더 **바로 안에** 복사해 넣습니다. (예: `deck.pptx`)

슬라이드 노트(발표자 노트)에 각 페이지 대본을 미리 적어두면 자동으로 읽어옵니다. 대본을 어떻게 써야 하는지는 아래 [대본 작성 방법](#대본-작성-방법)을 꼭 먼저 읽어주세요.

### 7단계. 실행!

```bash
venv/bin/python ppt2video.py --pptx deck.pptx --out output.mp4
```

별도 대본 파일(`script.json`)을 쓰는 경우:

```bash
venv/bin/python ppt2video.py --pptx deck.pptx --script script.json --out output.mp4
```

PDF를 이미 직접 만들어뒀다면 `--pdf`로 지정하세요. PPTX → PDF 자동 변환 과정을 건너뛰기 때문에 더 빠르게 실행됩니다:

```bash
venv/bin/python ppt2video.py --pdf deck.pdf --script script.json --out output.mp4
```

> **파일을 못 찾는다는 오류가 나면**: `--pptx`/`--pdf`로 준 파일이 지금 명령어를 실행하는 폴더 기준 경로에 없다는 뜻입니다.
> 오류 메시지에 찍히는 절대경로를 보고 파일을 그 위치로 옮기거나, 경로를 정확히 적어주세요 (예: `sample/deck.pptx`).

### 8단계 (선택). 먼저 샘플로 테스트해보기

```bash
venv/bin/python sample/make_test_pdf.py
venv/bin/python ppt2video.py --pdf sample/test.pdf --script sample/script.json --out sample/test_output.mp4
```

`sample/test_output.mp4`가 생기면 정상 작동하는 것입니다.

---

## 대본 작성 방법

대본은 두 가지 방법 중 하나로 준비할 수 있습니다.

**방법 1. PPT 슬라이드 노트에 적기 (추천)**

PowerPoint에서 슬라이드 아래쪽 "슬라이드 노트(발표자 노트)" 칸에 그 슬라이드에서 읽을 내용을 그대로 적어두면, 실행할 때 `--pptx`만 지정해도 자동으로 읽어옵니다. 노트가 없는 슬라이드(제목 슬라이드 등)는 자동으로 무음 처리됩니다.

**방법 2. 별도 파일(`script.json`)로 작성**

PPT 노트를 쓰지 않고 싶다면 `script.json` 파일을 만들어서 슬라이드 순서대로 대본을 적습니다.

```json
[
  "1페이지에서 하고 싶은 말",
  "2페이지에서 하고 싶은 말"
]
```

> **대본은 반드시 한국어로만 쓰세요.**
> 영어 단어나 전문용어를 원문 그대로 섞어 쓰면 한국어 TTS가 발음을 이상하게 뭉개서 읽습니다.
> 영어 용어는 **한글 발음 그대로 표기**하세요.
> 예) `Transformer` → `트랜스포머`, `Attention` → `어텐션`, `Feed-Forward Network` → `피드포워드 네트워크`, `Encoder-Decoder` → `인코더-디코더`

---

## 목소리 바꾸기

기본 목소리는 `M1`입니다. 다른 목소리로 바꾸려면 `--voice` 옵션을 추가하세요.

```
ppt2video.py --pptx deck.pptx --out output.mp4 --voice F2
```

사용 가능한 목소리 목록 보기:

```
ppt2video.py --list-voices
```

(`M1~M5`: 남성 목소리 5종, `F1~F5`: 여성 목소리 5종. 실행할 때는 위처럼 앞에 `venv\Scripts\python`(Windows) 또는 `venv/bin/python`(macOS/Linux)을 붙이세요.)

강의용으로 테스트했을 때 `M3`/`M4`가 비교적 안정적인 톤으로 들렸습니다. 강의 영상은 내용이 길기 때문에 기본 속도(`1.05`)보다 조금 빠른 `1.2` 정도가 덜 늘어지고 듣기 편했습니다.

```powershell
venv\Scripts\python ppt2video.py --pptx deck.pptx --out output.mp4 --voice M3 --speed 1.2
```

목소리는 취향 차이가 크니, 아래 스크립트로 10개 목소리를 전부 짧게 만들어서 직접 들어보고 고르는 걸 추천합니다.

```powershell
venv\Scripts\python sample\make_voice_samples.py
```

`sample\voices\` 폴더에 `M1.wav` ~ `F5.wav` 10개 파일이 생성됩니다.

<details>
<summary>참고: 목소리별 음높이 실측값</summary>

같은 문장을 10개 목소리로 만들어서 평균 음높이(F0)를 실제로 측정한 결과입니다. 숫자가 낮을수록 저음(깊은 목소리), 높을수록 고음입니다.

| 목소리 | 평균 음높이 | 특성 |
|---|---|---|
| `M5` | 약 93Hz | 남성 중 가장 저음. 무게감 있지만 다소 부담스러울 수 있음 |
| `M2` | 약 102Hz | 저음, 차분하고 진중한 톤 |
| `M3` | 약 108Hz | 중저음, 또렷하고 안정적 |
| `M4` | 약 117Hz | 중저음, `M3`보다 살짝 밝은 톤 |
| `M1` | 약 143Hz | 남성 중 가장 밝은 톤, 젊은 느낌 |
| `F5` | 약 167Hz | 여성 중 가장 낮은 톤, 차분함 |
| `F3` | 약 180Hz | 중간 톤 |
| `F4` | 약 187Hz | 중간 톤, 또렷함 |
| `F1` | 약 199Hz | 밝은 톤 |
| `F2` | 약 217Hz | 여성 중 가장 밝고 높은 톤 |

</details>

---

## 주요 옵션

| 옵션 | 설명 | 기본값 |
|---|---|---|
| `--pdf` | 슬라이드 이미지 소스 PDF. 없으면 `--pptx`로 자동 변환 | - |
| `--pptx` | 원본 PPTX. 노트를 대본으로 자동 추출하거나, `--pdf` 없을 때 PDF 자동 변환용으로 사용 | - |
| `--script` | 대본 파일 (.json / .txt) | - |
| `--out` | 출력 영상 경로 | `output.mp4` |
| `--lang` | TTS 언어 코드 | `ko` |
| `--voice` | Supertonic 내장 목소리 (`M1~M5`, `F1~F5`) | `M1` |
| `--speed` | TTS 발화 속도 배율 | `1.05` |
| `--steps` | 합성 스텝 수. 높을수록 음질 좋지만 느림 | `8` |
| `--sentence-pause` | 한 페이지 안에서 문장과 문장 사이 쉬는 시간(초) | `0.4` |
| `--pad` | 각 페이지 음성 뒤 여백(초) | `0.4` |
| `--min-duration` | 대본 없는 페이지 노출 시간(초) | `1.2` |
| `--width` | 렌더링 슬라이드 가로 해상도 | `1920` |
| `--fps` | 출력 영상 fps | `25` |
| `--workdir` | 중간 산출물 폴더 | `build` |
| `--keep-temp` | 중간 산출물 보존 | 끔 |

## 동작 방식 (참고)

1. **슬라이드 이미지**: `--pdf`로 준 PDF의 각 페이지를 이미지로 렌더링합니다.
   `--pdf` 없이 `--pptx`만 주면 자동으로 PDF를 만들어 사용합니다 (Windows: PowerPoint, macOS/Linux: LibreOffice).
2. **페이지별 대본**: `--pptx`(슬라이드 노트 자동 추출) 또는 `--script`(별도 파일) 중 하나로 지정합니다.
   - `.json`: 문자열 배열, 예) `["1페이지 대본", "2페이지 대본"]`
   - `.txt`: 페이지 사이를 `===`로 구분 (없으면 빈 줄 두 번으로 구분)
   - PDF 페이지 수와 대본 개수가 다르면 경고를 띄우고 짧은 쪽에 맞춥니다. 대본이 빈 페이지는 무음으로 처리됩니다.
3. **TTS 음성 생성**: Supertonic(ONNX 기반 경량 오픈소스 TTS)으로 각 페이지 대본을 음성으로 합성합니다.
   CPU만으로도 빠르게 동작하고 GPU/PyTorch가 필요 없습니다. 문장 단위로 나눠서 따로 합성한 뒤
   그 사이에 짧은 무음(`--sentence-pause`)을 넣기 때문에 여러 문장을 붙여 읽어도 숨 쉬는 구간 없이
   뭉개지지 않습니다.
4. **영상 합성**: 이미지 + 음성으로 페이지별 클립을 만든 뒤 순서대로 이어붙여 최종 mp4를 만듭니다.

## 참고 / 주의사항

- 첫 실행 시 Supertonic 모델(수십MB, HuggingFace Hub)이 자동 다운로드되어 캐시됩니다. 그다음부터는 재다운로드하지 않습니다.
- CPU만으로도 충분히 빠르게 동작합니다.
