# Meta Ads 리포팅 대시보드

Meta(Facebook/Instagram) 광고 데이터를 가져와서 보기 좋게 정리해주는 웹 대시보드입니다.

**할 수 있는 것:**
- 광고 계정 여러 개 등록하고 전환해서 보기
- 주간 / 월간 인사이트 (전 기간 대비 변동 자동 분석)
- Target × Targeting × Creative 단계로 펼쳐서 보기
- 일자별 데이터 (전체 / 타겟별 / 소재별로 전환)
- Excel 파일로 다운로드 (5개 시트 자동 생성)

---

## 시작하기 (Windows 기준)

### 1단계 — Python 설치하기

이미 설치되어 있는지 확인부터:
1. Windows 시작 버튼 누르고 **"명령 프롬프트"** 검색해서 실행
2. 검은 창에 다음 입력 후 엔터:
   ```
   python --version
   ```
3. `Python 3.x.x` 같은 글자가 나오면 OK. 다음 단계로 넘어가세요.
4. 만약 "Python을 찾을 수 없습니다" 같은 메시지가 뜨거나 Microsoft Store 창이 열리면 → 직접 설치 필요:
   - https://www.python.org/downloads/ 접속
   - **Download Python 3.x.x** 큰 버튼 클릭
   - 다운로드된 설치 파일 실행
   - ⚠️ **중요:** 첫 화면 하단의 **"Add Python to PATH"** 체크박스 꼭 체크하고 **Install Now** 클릭
   - 설치 완료 후 명령 프롬프트 **닫았다가 다시 열고** 위 1~3단계 다시 확인

### 2단계 — 프로젝트 가져오기

이미 `git clone` 해서 폴더가 있다고 가정합니다. 명령 프롬프트에서 프로젝트 폴더로 이동:

```
cd C:\projects\JAY
```
(클론한 위치가 다르면 그 경로로)

### 3단계 — 필요한 패키지 설치하기

명령 프롬프트에서 다음 한 줄 입력하고 엔터:

```
pip install -r requirements.txt
```

설치 진행되면서 여러 줄이 출력되고 마지막에 `Successfully installed ...` 가 뜨면 완료. (1~2분 걸릴 수 있음)

만약 `pip` 명령을 찾을 수 없다고 나오면, Python 설치 시 PATH 추가가 안 된 거예요. 1단계 재설치 또는 `python -m pip install -r requirements.txt` 로 대체 시도.

### 4단계 — Meta 광고 API 토큰 발급받기

대시보드가 광고 데이터를 가져오려면 Meta에서 발급해주는 "토큰"이 필요해요.

1. https://developers.facebook.com/apps 접속 후 Facebook 로그인
2. **앱 만들기** 클릭 → 유형 **"비즈니스"** 선택 → 이름 입력 → 생성
3. 생성된 앱 대시보드에서 좌측 **"제품 추가" → "Marketing API" → 설정** 클릭
4. 좌측 **"앱 설정 → 기본"** 들어가서 **앱 시크릿** 옆의 **"표시"** 클릭 → 둘 다(앱 ID, 앱 시크릿) 메모장에 복사해두세요

5. https://developers.facebook.com/tools/explorer 접속
6. 우측 상단 **"Meta 앱"** 드롭다운에서 방금 만든 앱 선택
7. **"권한 추가"** 클릭 → **`ads_read`** 만 체크 (수정 권한이 필요하면 `ads_management`도)
8. **"액세스 토큰 생성"** → Facebook 권한 승인 → 짧은 토큰 발급됨 (1~2시간 유효)

이대로 두면 1~2시간 후 만료되니까 **60일짜리 토큰으로 교환** 합니다:

브라우저 주소창에 아래 URL을 붙여넣되, `{}` 부분을 본인 값으로 교체:
```
https://graph.facebook.com/v25.0/oauth/access_token?grant_type=fb_exchange_token&client_id={앱ID}&client_secret={앱시크릿}&fb_exchange_token={방금받은단기토큰}
```

엔터 누르면 JSON으로 응답이 나오는데, 그 안의 `access_token` 값이 **60일 유효한 새 토큰**. 복사해두세요.

### 5단계 — 광고 계정 ID 확인하기

1. https://adsmanager.facebook.com 접속
2. 좌측 상단 계정 선택 드롭다운에서 **숫자**만 복사 (예: `568276609290906`)
3. 앞에 `act_` 를 붙인 형태가 광고 계정 ID (예: `act_568276609290906`)

### 6단계 — 토큰 / 계정 정보 입력하기

프로젝트 폴더 안의 **`.env`** 파일을 메모장으로 열어요. (파일이 없으면 `.env.example` 복사해서 `.env`로 이름 변경)

내용을 이렇게 채우세요:
```
META_API_VERSION=v25.0
META_ACCESS_TOKEN=여기에_4단계에서_받은_60일_토큰_붙여넣기
META_AD_ACCOUNT_IDS=act_여기에_5단계_계정_ID
```

여러 계정 등록하려면 콤마로 구분: `META_AD_ACCOUNT_IDS=act_111,act_222,act_333`

저장하고 닫기.

### 7단계 — 서버 실행하기

명령 프롬프트에서 (프로젝트 폴더에 있는 상태):

```
python app.py
```

다음과 같은 메시지가 뜨면 성공:
```
 * Serving Flask app 'app'
 * Debug mode: on
 * Running on http://127.0.0.1:5000
```

⚠️ **이 명령 프롬프트 창은 닫지 마세요.** 닫으면 서버가 꺼져요.

### 8단계 — 브라우저로 접속하기

브라우저(Chrome/Edge/Safari 아무거나) 주소창에:
```
http://127.0.0.1:5000
```

대시보드가 뜨면 성공!

---

## 사용 방법

### 기본 흐름
1. 우측 상단에 **● 토큰 설정됨** 배지 확인 (안 뜨면 .env 다시 확인)
2. **광고 계정** 카드에서 등록된 계정 선택
3. **기간** 선택 (기본: 최근 7일) → **조회** 클릭
4. 데이터가 로드되면 아래 카드들이 표시됩니다:
   - **📊 인사이트 비교** — 주간/월간 탭, 전 기간 대비 변동 자동 분석
   - **📊 SUMMARY** — Target × Targeting × Creative 계층 (Targeting 행 클릭 → 펼침)
   - **📅 일자별 상세** — 날짜 × Target/Targeting/Creative (컨텍스트 바로 전환)
   - **상세 데이터** — 광고 단위 일별 행, 컬럼 헤더에 필터
   - **🎯 타겟 × 에셋** — 캠페인별 진행률 / 소진률 막대

### 광고 계정 추가/삭제
- **+** 버튼 → 광고 계정 ID 입력 → 추가 (Meta API에서 즉시 검증)
- **✕** 버튼 → 현재 선택된 계정 목록에서 제거 (Meta 광고 계정 자체는 안 건드림)

### 필터 사용
SUMMARY 카드 상단의 필터 바에서:
- Target, Targeting, 캠페인, 광고세트, 광고 다중 선택 가능
- 적용하면 SUMMARY / 일자별 / 상세 데이터 / 인사이트 비교 전부 즉시 재집계

### 인사이트 비교
- **주간 탭**: 기본 = 가장 최근 완료된 주(월~일) vs 그 전 주
- **월간 탭**: 기본 = 지난 달 vs 전전 달
- 날짜 직접 수정 가능 → **조회** 버튼
- 하단 **🎯 메인 KPI** 선택자에서 노출/클릭/뷰/전환 중 하나 선택 → 자동 인사이트가 그 KPI 중심으로 재작성

### Excel 다운로드
**📥 Excel 다운로드** 버튼 → 5개 시트 들어있는 파일 다운로드:
1. 주간 인사이트
2. SUMMARY (Target × Targeting × Creative)
3. 일자별_Adults
4. 일자별_Young Learners
5. 일자별_Kinder

---

## 토큰 만료 대응

토큰은 **60일** 유효. 만료되면 대시보드에 데이터가 안 뜨고 에러 메시지가 표시돼요.

해결:
1. 4단계 다시 (단기 토큰 발급 → 60일 토큰 교환)
2. `.env` 파일의 `META_ACCESS_TOKEN` 값만 새 토큰으로 교체
3. 서버 재시작 (명령 프롬프트에서 `Ctrl+C` 누른 다음 `python app.py` 다시 실행)

---

## 문제 해결

### "Python을 찾을 수 없습니다" / "pip를 찾을 수 없습니다"
- Python 설치 시 **"Add Python to PATH"** 체크 안 한 게 원인
- Python을 다시 설치하거나, 명령 앞에 `python -m`을 붙여 시도: `python -m pip install -r requirements.txt`

### "포트 5000이 이미 사용 중"
- 이미 다른 곳에서 서버가 떠 있는 거예요
- 모든 명령 프롬프트 창 닫고 다시 시도
- 또는 `app.py` 마지막 줄의 `port=5000` 을 `port=5001` 등으로 변경 후 그에 맞춰 브라우저 주소도 `http://127.0.0.1:5001`

### "OAuth Exception" / "Session has expired"
- 토큰이 만료됨 → 위 "토큰 만료 대응" 섹션 참고

### 광고 계정 추가했는데 "허용되지 않은 ID"
- 그 계정에 본인 토큰이 접근 권한이 없는 거예요
- Business Manager에서 본인을 그 광고 계정의 관리자로 추가했는지 확인

### Excel 다운로드 시 에러
- 화면에 에러 메시지(traceback)가 표시되니 캡쳐 떠두기
- 서버를 실행한 명령 프롬프트 창에도 자세한 에러 로그 출력됨

### 데이터가 안 뜸 / "data가 비어 있습니다"
- 선택한 기간에 광고 활동이 실제로 없거나, 토큰 권한 부족
- Meta Ads Manager에서 직접 같은 기간을 봐서 데이터가 있는지 비교

---

## 폴더 구조

```
JAY/
├── app.py              ← 서버 본체 (수정하지 마세요)
├── templates/
│   └── index.html      ← 화면 (수정하지 마세요)
├── requirements.txt    ← 필요한 패키지 목록
├── .env                ← 토큰/계정 ID (직접 채워야 하는 파일)
├── .env.example        ← .env 작성 예시
├── accounts.json       ← 등록한 광고 계정 목록 (자동 생성됨)
├── README.md           ← 이 문서
└── CLAUDE.md           ← 개발자/AI용 컨텍스트 문서
```

`.env`와 `accounts.json`은 본인 계정 정보가 들어있으니 **다른 사람한테 공유하면 안 됩니다.** `.gitignore`에 들어 있어서 git push 해도 자동으로 제외됩니다.

---

## 도움 요청

- 이 도구는 영국문화원(British Council) Korea 광고 운영용으로 만들어진 베타 버전이에요
- 기능 추가/수정이 필요하면 Claude AI 어시스턴트와 함께 작업한 컨텍스트가 `CLAUDE.md` 에 정리되어 있으니 그걸 참고해서 개선 요청하면 됩니다
- 자세한 작업 히스토리와 설계 결정은 `CLAUDE.md` 참고
