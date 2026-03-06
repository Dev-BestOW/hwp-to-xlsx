# Excel → 한글(HWPX) 자동 입력

엑셀 데이터를 한글(HWPX) 템플릿에 자동으로 채워넣어 개별 파일을 생성하는 웹 앱입니다.

## 배포 URL

https://hwp-to-xlsx-4qgsfq48igyuedtlvns2d6.streamlit.app/

## 주요 기능

- **한글 템플릿 업로드**: `{{매도인}}`, `{{매수인}}` 같은 플레이스홀더가 포함된 HWPX 파일 업로드 (여러 개 가능)
- **엑셀 파일 업로드**: 플레이스홀더와 매칭되는 컬럼명을 가진 엑셀 데이터 업로드
- **웹에서 직접 입력**: 엑셀 없이 웹 UI에서 스프레드시트처럼 직접 데이터 입력 가능
- **자동 생성**: 행 수 x 템플릿 수만큼 개별 HWPX 파일 생성
- **일괄 다운로드**: ZIP 파일로 한번에 다운로드

## 사용 방법

1. 한글에서 바꿀 부분을 `{{플레이스홀더}}` 형식으로 작성 후 HWPX로 저장
2. 웹 앱에서 템플릿 업로드
3. 엑셀 업로드 또는 웹에서 직접 데이터 입력
4. 파일 생성 버튼 클릭 후 다운로드

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

## CLI 사용

```bash
python main.py --template template.hwpx --data data.xlsx --output ./output
```

## 기술 스택

- Python 3.11
- Streamlit
- openpyxl, pandas, lxml
- HWPX 직접 ZIP/XML 처리
