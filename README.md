# CS 결과보고 AI 자동작성기

CS 주간 현황판 + 개인 2분기 목표 파일을 업로드하면  
AI가 OKR · KPI 포함 결과보고서를 자동으로 작성하고 엑셀로 추출해주는 앱입니다.

## 배포 방법 (Streamlit Community Cloud)

### 1. GitHub 레포 생성
```
git init
git add .
git commit -m "init"
git remote add origin https://github.com/[아이디]/cs-report-app.git
git push -u origin main
```

### 2. Streamlit Cloud 설정
1. https://share.streamlit.io 접속
2. "New app" → GitHub 레포 선택
3. Main file: `app.py`
4. **Secrets 설정** (중요!)
   - Advanced settings → Secrets 탭
   - 아래 내용 붙여넣기:
   ```toml
   ANTHROPIC_API_KEY = "sk-ant-..."
   ```

### 3. 배포 완료
`https://[아이디]-cs-report-app-app-xxxx.streamlit.app` 형태 URL 발급

## 파일 구조
```
cs_report_app/
├── app.py                    # 메인 앱
├── requirements.txt          # 패키지 목록
└── .streamlit/
    ├── config.toml           # 테마 설정
    └── secrets.toml          # 로컬 테스트용 (GitHub에 올리지 마세요!)
```

## 주요 기능
- CS 주간 현황판 Excel 파싱 (2분기 전체)
- 개인 목표 Excel 파싱 (목표/등급/미션 추출)
- Claude AI 자동 분석 (OKR · KPI · 날짜 · 상호 · 과정 · 결과)
- 엑셀 출력: OKR요약 / KPI실적 / 목표별상세 / CS주간요약 / 종합평가
- 원본 목표 양식에 결과 직접 주입 옵션
- JSON 데이터 다운로드

## 주의사항
- `.streamlit/secrets.toml` 파일은 `.gitignore`에 추가하세요
- API Key는 Streamlit Cloud Secrets에만 저장하세요
