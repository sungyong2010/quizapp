# 📚 Daily Vocabulary Quiz App (Google Sheets 연동)

이 프로젝트는 Google Sheets API를 활용하여 실시간으로 퀴즈 데이터를 불러오고, 사용자가 한글 단어에 대한 영어 정답을 입력해야 하는 전체 화면 퀴즈 프로그램입니다. 학습용 키오스크 환경이나 어린이 교육용으로 적합합니다.

---

## 🧩 주요 기능

- Google Sheets에서 실시간 퀴즈 데이터 불러오기
- 날짜별 시트 자동 선택 (`YYYY-MM-DD` 형식)
- 전체 화면 GUI (tkinter 기반)
- 정답을 맞춰야 다음 문제로 진행
- Alt+F4 차단으로 종료 방지
- 엔터키로도 정답 제출 가능

---

## 📄 Google Sheets 퀴즈 데이터 형식

| 한글 단어 | 영어 정답 |
|-----------|------------|
| 필요하다   | need       |
| 돈        | money      |
| 관심       | interest   |
| ...       | ...        |

- 시트 이름은 날짜 형식: `2025-10-26`, `2025-10-27` 등
- 예시 문서: [퀴즈 데이터 시트](https://docs.google.com/spreadsheets/d/1BHkAT3j75_jq5qM5p1AZ73NaR4JhcxP7uBeWZRE0CD8/edit?usp=sharing)

---

## 🛠️ 설치 및 실행 방법

### 1. 필수 패키지 설치
```bash
pip install gspread oauth2client tkinter
