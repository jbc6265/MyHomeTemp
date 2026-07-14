# MyHomeTemp

LG ThinQ API와 Google Apps Script를 이용해 집의 온도/습도 데이터를 Google Sheets에 기록하고, Apps Script 웹앱 HTML 대시보드로 확인하는 프로젝트입니다.

## 구성

- `Code.js`: LG ThinQ 상태 조회, Google Sheets 기록, 웹앱 데이터 API
- `Index.html`: HTML 대시보드 화면
- `appsscript.json`: Apps Script 프로젝트 설정

## 로컬 동기화

이 프로젝트는 `clasp`로 Apps Script와 동기화합니다.

```bash
npx @google/clasp push
```

민감하지는 않지만 프로젝트 식별자가 들어 있는 `.clasp.json`은 저장소에 올리지 않습니다.
