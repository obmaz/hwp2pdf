# HWP-to-PDF-Batch-Converter

[![Build EXE](https://github.com/obmaz/hwp2pdf/actions/workflows/build.yml/badge.svg)](https://github.com/obmaz/hwp2pdf/actions/workflows/build.yml)
![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)

본 프로젝트는 한컴오피스 HWP 및 HWPX 문서를 고속으로 PDF로 일괄 변환하는 Windows 전용 데스크탑 애플리케이션입니다. Python의 `win32com` 라이브러리를 통해 한컴오피스 Automation API를 제어하며, 비동기 멀티스레딩 아키텍처를 채택하여 대량의 문서 처리 시에도 안정적인 성능과 UI 응답성을 제공합니다.

---

## 🛠 핵심 기술 스택 및 아키텍처 (Technical Stack)

- **Language**: Python 3.8+
- **GUI Framework**: Tkinter / CustomTkinter (현대적인 UI 컴포넌트)
- **Runtime Automation**: `pywin32` (Microsoft COM/OLE Automation)
- **Handling UI Events**: `tkinterdnd2` (Native Windows Drag & Drop Integration)
- **Concurrency**: `threading` 모듈을 이용한 Worker/UI 스레드 분리
- **Build System**: PyInstaller & GitHub Actions (CI/CD 자동 빌드 지원)

---

## 🚀 주요 기능 (Key Features)

### 1. 지능형 파일 패키징 및 폴더 구조 유지
- **Recursive Scan**: 입력된 폴더 내의 모든 하위 디렉토리를 재귀적으로 탐색하여 HWP/HWPX 파일을 추출합니다.
- **Path Mapping**: 원본 디렉토리 구조를 분석하여 출력 경로에 동일한 계층 구조를 자동 생성하므로 결과물 관리가 용이합니다.

### 2. 고성능 백그라운드 프로세싱 및 최적화
- **Multi-Threading**: 변환 로직을 별도의 워커 스레드에서 실행하여 대량 변환 중에도 GUI 프리징 현상을 완벽히 차단합니다.
- **Automation Stability**: 
  - `Visible = False`: 한글 프로세스를 백그라운드 은닉 모드로 실행하여 리소스 소모를 최소화합니다.
  - `SetMessageBoxMode(0x20)`: 폰트 누락, 손상된 문서 등 작업 중 발생하는 모든 차단형 팝업(Blockage)을 무시하도록 설정하여 자동화 무결성을 보장합니다.

### 3. 직관적인 사용자 인터페이스
- **DND Support**: 시스템 탐색기로부터 파일 및 폴더의 드래그 앤 드롭을 지원합니다.
- **Real-time Status Tracking**: 각 개별 파일의 변환 상태(Pending, Processing, Success, Failure)를 데이터 그리드에 실시간으로 업데이트합니다.

---

## 📋 시스템 요구 사항 (System Requirements)

- **OS**: Windows 10 / 11
- **Software**: **한컴오피스 한글** 설치 필수 (Automation API 서버 제공 필요)
- **Runtime**: Python 3.8 이상 (소스 실행 시)

---

## 🔧 설치 및 실행 가이드 (Setup & Execution)

### 의존성 설치
환경에 필요한 라이브러리를 설치합니다.
```bash
pip install -r requirements.txt
```

### 애플리케이션 실행
```bash
python src/main.py
```

---

## 📦 빌드 및 배포 (Build & Deployment)

### 로컬 빌드
PyInstaller를 사용하여 단일 실행 파일(`.exe`)을 생성할 수 있습니다.
```bash
python -m PyInstaller HWP_to_PDF.spec
```
또는 `build_exe.bat` 파일을 실행하여 간편하게 빌드 프로세스를 완료할 수 있습니다.

### CI/CD (GitHub Actions)
본 저장소에 `push` 시 GitHub Actions를 통해 Windows 환경에서 실행 파일이 자동으로 빌드됩니다. 빌드된 결과물은 각 Actions 실행의 **Artifacts** 탭에서 다운로드 가능합니다.

---

## 📎 Troubleshooting

- **COM 인터페이스 오류**: 한컴오피스가 정상적으로 설치되지 않았거나 레지스트리에 COM 서버가 등록되지 않은 경우 발생할 수 있습니다. 한글 프로그램을 1회 실행 후 시도해 보십시오.
- **권한 문제**: 관리자 권한으로 실행 중인 앱과 일반 권한 앱 사이에는 Windows 보안 정책상 드래그 앤 드롭이 제한될 수 있습니다.
- **변환 실패**: 암호가 걸려 있거나 파일이 크게 손상된 경우 매크로 실행이 중단될 수 있습니다.

---

## 📄 License
[MIT License](LICENSE) (또는 해당 라이선스 명시)
