# Streamlit Cloud 배포 가이드

이 프로젝트는 로컬 실행용으로도 동작하지만, Streamlit Cloud에 올려 공개 링크로 공유하려면 몇 가지 준비가 필요합니다.

## 1. 왜 localhost 링크는 다른 사람이 열 수 없나?

localhost는 사용자의 개인 PC 안에서만 열 수 있는 주소입니다. 다른 사람이 접속하려면 GitHub 저장소에 업로드하고 Streamlit Cloud 같은 호스팅 서비스에 배포해야 합니다.

## 2. GitHub에 올릴 때 .env를 올리면 안 되는 이유

.env 파일에는 Gemini API 키와 같은 민감 정보가 들어 있을 수 있습니다. 공개 저장소에 올리면 키가 유출될 수 있으므로, .env는 GitHub에 올리지 않고 Streamlit Cloud의 Secrets로 관리해야 합니다.

## 3. Streamlit Cloud에서 배포하는 순서

1. GitHub 저장소를 생성합니다.
2. 이 프로젝트 폴더의 파일들을 GitHub에 업로드합니다.
3. Streamlit Cloud에 로그인합니다.
4. 새 앱을 생성합니다.
5. Repository를 선택합니다.
6. Branch는 main으로 설정합니다.
7. Main file path는 app.py로 설정합니다.
8. Deploy 버튼을 눌러 배포합니다.

## 4. Secrets 설정

Streamlit Cloud 앱 설정에서 Secrets에 다음 항목을 추가합니다.

```toml
[GEMINI_API_KEY]
```

값으로 실제 Gemini API 키를 넣습니다.

## 5. 배포 후 확인할 항목

- 앱이 정상적으로 열리는지
- Vector DB가 로드되는지
- 질문 시나리오 테스트 모드가 보이는지
- Gemini API 키가 없을 때 앱이 오류 없이 안내를 표시하는지

## 6. 앱이 안 열릴 때 확인할 로그 위치

Streamlit Cloud의 앱 화면에서 "Manage app" 또는 "Logs"를 확인하면 배포 로그를 볼 수 있습니다. 오류가 있으면 여기서 원인을 확인하세요.

## 7. 배포 전 꼭 확인할 것

- .env는 GitHub에 포함하지 않았는지
- requirements.txt에 필요한 패키지가 모두 들어 있는지
- app.py가 문법 오류 없이 실행되는지
