# Tricare Streamlit 배포 가이드

이 프로젝트의 Streamlit 앱을 Windows용 단일 실행 파일(`TricareApp.exe`)로 패키징해, 다른 PC에서 파이썬 없이 실행하는 방법을 정리했습니다.

## 구성 요소
- `app.py`: Streamlit 메인 앱
- `processor.py`: 비즈니스 로직 모듈
- `launch.py`: PyInstaller 엔트리포인트 (포트/브라우저 설정 포함)
- `TricareApp.spec`: 빌드 설정 (Streamlit 정적 자산 및 스크립트 포함)

## 사전 준비 (빌드 머신)
1) Python 3.10+ 설치  
2) 의존성 설치
```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller
```

## 빌드 방법
프로젝트 루트(`TricareApp.spec`가 있는 곳)에서 실행합니다.
```powershell
# 선택: 이전 빌드 산출물 정리
Remove-Item -Recurse -Force build, dist -ErrorAction SilentlyContinue

# 빌드
pyinstaller TricareApp.spec
```
빌드 후 `dist\TricareApp.exe`가 생성됩니다.

### 콘솔 숨기기(릴리스 용)
최종 배포 시 콘솔을 숨기고 싶다면 스펙 대신 직접 명령으로 빌드합니다.
```powershell
pyinstaller --noconfirm --onefile --noconsole `
  --collect-all streamlit `
  --name TricareApp launch.py
```

## 배포/실행 (타겟 PC)
1) `dist\TricareApp.exe`를 원하는 폴더로 복사  
2) 필요한 데이터 폴더/파일이 있다면 동일 위치에 함께 배치  
3) 더블 클릭 실행 → 기본 브라우저가 자동으로 열리며 기본 포트는 `8501`  
   - 브라우저가 자동으로 안 열리면 직접 `http://localhost:8501` 접속

## 트러블슈팅
- **포트 충돌**: 다른 프로세스가 8501 사용 중이면 종료하거나 `launch.py`에서 `--server.port` 값을 바꾼 뒤 다시 빌드.  
- **모듈 누락 오류**: `ModuleNotFoundError` 발생 시 해당 모듈을 `requirements.txt`에 추가 후 `pip install -r requirements.txt` 실행 후 재빌드.  
- **Streamlit 경고**: `server.enableCORS` 관련 경고는 기본 설정에서 무시 가능.  
- **빌드 실패**: `__file__` 관련 에러는 `TricareApp.spec` 최신 버전을 사용해 재빌드.

## 참고
- 실행 시 `STREAMLIT_GLOBAL_DEVELOPMENT_MODE`를 자동으로 끄고, `8501` 포트로 고정해 패키징되었습니다.
- 필요 시 `data/` 등 리소스를 exe와 같은 폴더에 두거나 `--add-data` 옵션으로 포함하세요.

