# ppt2video

PPT/PDF 슬라이드 + 페이지별 대본을 오픈소스 TTS([Supertonic](https://pypi.org/project/supertonic/))로
음성 합성해서 자동으로 나레이션 영상(mp4)을 만들어주는 CLI 도구입니다.

---

## 실행하는 법 (처음부터 순서대로)

### 0단계. 준비물

- Windows PC
- PowerShell (윈도우에 기본 내장, 시작 메뉴에서 "PowerShell" 검색)
- 인터넷 연결 (처음 한 번은 프로그램/모델 다운로드 필요)

### 1단계. 이 저장소를 내 컴퓨터로 받기

PowerShell을 열고 아래 명령어를 순서대로 입력합니다. (원하는 폴더로 이동한 뒤 실행하세요. 예: `cd C:\`)

```powershell
git clone https://github.com/jhlee001010/SlideVoice.git
cd SlideVoice
```

> git이 없다면 GitHub 페이지에서 "Code > Download ZIP"으로 받아서 압축을 풀어도 됩니다.

### 2단계. Python 설치 확인

```powershell
python --version
```

`Python 3.x.x` 같은 게 안 뜨고 오류가 나면, 아래 명령으로 설치합니다.

```powershell
winget install -e --id Python.Python.3.12
```

설치 후에는 **PowerShell 창을 새로 열어야** 인식됩니다.

### 3단계. ffmpeg 설치 확인

```powershell
ffmpeg -version
```

오류가 나면 설치합니다.

```powershell
winget install -e --id Gyan.FFmpeg
```

이것도 설치 후 **새 PowerShell 창**에서 진행하세요.

### 4단계. 가상환경 만들고 필요한 패키지 설치

(SlideVoice 폴더 안에서 실행)

```powershell
python -m venv venv
venv\Scripts\pip install -r requirements.txt
```

몇 분 정도 걸릴 수 있습니다.

### 5단계. 영상으로 만들 슬라이드 준비

만든 `.pptx` 파일을 `SlideVoice` 폴더 안에 복사해 넣습니다. (예: `deck.pptx`)

- **PDF는 따로 안 만들어도 됩니다.** `--pptx`만 주면 설치된 PowerPoint로 자동으로 PDF를 만들어서 씁니다.
  (Windows + PowerPoint 설치가 되어 있어야 합니다. 둘 다 이 PC에는 이미 있습니다.)
- 슬라이드 노트(발표자 노트)에 각 페이지 대본을 미리 적어두면 자동으로 읽어옵니다.
- 노트 대신 대본을 직접 텍스트로 넣고 싶다면 6단계 참고.
- PDF를 직접 준비하고 싶다면 PowerPoint에서 **파일 > 내보내기 > PDF/XPS 만들기**로 저장한 뒤 `--pdf`로 지정해도 됩니다 (변환 시간을 아낄 수 있음).

### 6단계 (선택). 대본을 별도 파일로 직접 쓰고 싶다면

PPTX 노트를 안 쓰고 싶으면 `script.json` 같은 파일을 만들어서 페이지 순서대로 대본을 적습니다.

```json
[
  "1페이지에서 하고 싶은 말",
  "2페이지에서 하고 싶은 말"
]
```

### 7단계. 실행!

PPTX 노트를 대본으로 쓰는 경우 (PDF 자동 변환):

```powershell
venv\Scripts\python ppt2video.py --pptx deck.pptx --out output.mp4
```

직접 쓴 대본 파일(`script.json`)을 쓰는 경우:

```powershell
venv\Scripts\python ppt2video.py --pptx deck.pptx --script script.json --out output.mp4
```

PDF를 이미 직접 만들어뒀다면 `--pdf deck.pdf`를 추가로 지정하면 변환을 건너뛰고 더 빠르게 실행됩니다.

실행하는 동안 화면에 `[1/4] ... [2/4] ... [3/4] ... [4/4] ...` 진행 상황이 나오고,
끝나면 같은 폴더에 `output.mp4`가 생성됩니다.

### 8단계 (선택). 먼저 샘플로 테스트해보기

내 슬라이드 없이 미리 작동하는지만 확인하고 싶다면:

```powershell
venv\Scripts\python sample\make_test_pdf.py
venv\Scripts\python ppt2video.py --pdf sample\test.pdf --script sample\script.json --out sample\test_output.mp4
```

`sample\test_output.mp4`가 생기면 정상 작동하는 것입니다.

---

## 목소리 바꾸기

기본 목소리는 `M1`입니다. 다른 목소리로 바꾸려면 `--voice` 옵션을 추가하세요.

```powershell
venv\Scripts\python ppt2video.py --pdf deck.pdf --pptx deck.pptx --out output.mp4 --voice F2
```

사용 가능한 목소리 목록 보기:

```powershell
venv\Scripts\python ppt2video.py --list-voices
```

(`M1`~`M5`: 남성 목소리 5종, `F1`~`F5`: 여성 목소리 5종)

---

## 동작 방식 (참고)

1. **슬라이드 이미지**: `--pdf`로 준 PDF의 각 페이지를 이미지로 렌더링합니다.
   `--pdf` 없이 `--pptx`만 주면 설치된 PowerPoint로 자동으로 PDF를 만들어 사용합니다 (Windows + PowerPoint 필요).
2. **페이지별 대본**: `--pptx`(슬라이드 노트 자동 추출) 또는 `--script`(별도 파일) 중 하나로 지정합니다.
   - `.json`: 문자열 배열, 예) `["1페이지 대본", "2페이지 대본"]`
   - `.txt`: 페이지 사이를 `===`로 구분 (없으면 빈 줄 두 번으로 구분)
   - PDF 페이지 수와 대본 개수가 다르면 경고를 띄우고 짧은 쪽에 맞춥니다. 대본이 빈 페이지는 무음으로 처리됩니다.
3. **TTS 음성 생성**: Supertonic(ONNX 기반 경량 오픈소스 TTS)으로 각 페이지 대본을 음성으로 합성합니다.
   CPU만으로도 빠르게 동작하고 GPU/PyTorch가 필요 없습니다.
4. **영상 합성**: 이미지 + 음성으로 페이지별 클립을 만든 뒤 순서대로 이어붙여 최종 mp4를 만듭니다.

## 주요 옵션

| 옵션 | 설명 | 기본값 |
|---|---|---|
| `--pdf` | 슬라이드 이미지 소스 PDF. 없으면 `--pptx`로 자동 변환 | - |
| `--pptx` | 원본 PPTX. 노트를 대본으로 자동 추출하거나, `--pdf` 없을 때 PDF 자동 변환용으로 사용 | - |
| `--script` | 대본 파일 (.json / .txt) | - |
| `--out` | 출력 영상 경로 | `output.mp4` |
| `--lang` | TTS 언어 코드 | `ko` |
| `--voice` | Supertonic 내장 목소리 (`M1`~`M5`, `F1`~`F5`) | `M1` |
| `--speed` | TTS 발화 속도 배율 | `1.05` |
| `--steps` | 합성 스텝 수. 높을수록 음질 좋지만 느림 | `8` |
| `--pad` | 각 페이지 음성 뒤 여백(초) | `0.4` |
| `--min-duration` | 대본 없는 페이지 노출 시간(초) | `1.2` |
| `--width` | 렌더링 슬라이드 가로 해상도 | `1920` |
| `--fps` | 출력 영상 fps | `25` |
| `--workdir` | 중간 산출물 폴더 | `build` |
| `--keep-temp` | 중간 산출물 보존 | 끔 |

## 참고 / 주의사항

- 첫 실행 시 Supertonic 모델(수십MB, HuggingFace Hub)이 자동 다운로드되어 캐시됩니다. 그다음부터는 재다운로드하지 않습니다.
- CPU만으로도 충분히 빠르게 동작합니다.
