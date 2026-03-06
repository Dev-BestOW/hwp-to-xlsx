"""Excel → HWPX 자동 입력 웹 앱."""

import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

from src.excel_reader import read_excel
from src.hwpx_engine import replace_placeholders, find_placeholders_in_template
from src.utils import generate_output_filename


st.set_page_config(
    page_title="Excel → 한글 자동 입력",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Excel → 한글(HWPX) 자동 입력")
st.markdown("엑셀 데이터를 여러 한글 템플릿에 자동으로 채워넣어 개별 파일을 생성합니다.")

# --- 1. 템플릿 업로드 ---
st.header("1단계: 한글 템플릿 업로드")
st.markdown(
    "한글 파일에서 바꿀 부분을 `{{매도인}}`, `{{매수인}}` 처럼 이중 중괄호로 표시한 뒤 "
    "**HWPX 형식**으로 저장해주세요. **여러 개** 업로드 가능합니다."
)
template_files = st.file_uploader(
    "HWPX 템플릿 파일 (여러 개 가능)",
    type=["hwpx"],
    accept_multiple_files=True,
    help="한글에서 '다른 이름으로 저장' → HWPX 형식 선택",
)

# --- 2. 엑셀 업로드 ---
st.header("2단계: 엑셀 파일 업로드")
st.markdown(
    "첫 번째 행(헤더)에 템플릿의 플레이스홀더 이름과 동일한 컬럼명을 넣어주세요."
)
data_file = st.file_uploader(
    "엑셀 데이터 파일",
    type=["xlsx", "xls"],
    help="첫 번째 행이 헤더, 나머지가 데이터",
)

# --- 3. 옵션 ---
st.header("3단계: 옵션 설정")
col1, col2 = st.columns(2)
with col1:
    filename_pattern = st.text_input(
        "파일명 패턴",
        value="{매도인}",
        help="엑셀 컬럼명을 중괄호로 감싸서 사용 (예: {매도인})",
    )
with col2:
    st.markdown("&nbsp;")  # spacer
    st.info("예시: `{매도인}` → `홍길동_토지거래계약 허가 신청서.hwpx`\n\n템플릿 이름이 자동으로 붙습니다.")

# --- 4. 실행 ---
if template_files and data_file:
    st.divider()

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # 업로드된 템플릿 저장
        template_paths = []
        for tf in template_files:
            tp = tmpdir / tf.name
            tp.write_bytes(tf.getvalue())
            template_paths.append(tp)

        # 엑셀 저장 & 읽기
        data_path = tmpdir / data_file.name
        data_path.write_bytes(data_file.getvalue())

        try:
            rows = read_excel(data_path)
        except (FileNotFoundError, ValueError) as e:
            st.error(f"엑셀 오류: {e}")
            st.stop()

        # 미리보기
        st.subheader("📋 미리보기")

        st.markdown(f"**엑셀 데이터:** {len(rows)}행")
        st.dataframe(rows, use_container_width=True)

        st.markdown(f"**템플릿:** {len(template_paths)}개")
        excel_headers = set(rows[0].keys()) if rows else set()

        for tp in template_paths:
            try:
                placeholders = find_placeholders_in_template(tp)
            except Exception as e:
                st.error(f"{tp.name}: {e}")
                continue

            matched = set(placeholders) & excel_headers
            unmatched = set(placeholders) - excel_headers
            status = "✅" if not unmatched else "⚠️"
            detail = f"플레이스홀더: {', '.join(placeholders)}" if placeholders else "플레이스홀더 없음"
            if unmatched:
                detail += f" | **매칭 안 됨: {', '.join(unmatched)}**"
            st.markdown(f"- {status} `{tp.name}` — {detail}")

        total_files = len(rows) * len(template_paths)
        st.markdown(f"**생성될 파일 수:** {len(rows)}행 × {len(template_paths)}템플릿 = **{total_files}개**")

        # 생성 버튼
        st.divider()
        if st.button("🚀 한글 파일 생성", type="primary", use_container_width=True):
            output_dir = tmpdir / "output"
            output_dir.mkdir()

            progress = st.progress(0, text="파일 생성 중...")
            generated_files = []
            errors = []
            count = 0

            for i, row in enumerate(rows):
                for tp in template_paths:
                    template_label = tp.stem  # 확장자 제외한 파일명
                    ext = ".hwpx"
                    base_name = generate_output_filename(row, filename_pattern, "", i)
                    filename = f"{base_name}_{template_label}{ext}"

                    output_path = output_dir / filename

                    try:
                        replaced = replace_placeholders(tp, output_path, row)
                        generated_files.append((filename, output_path))
                    except Exception as e:
                        errors.append((filename, str(e)))

                    count += 1
                    progress.progress(count / total_files, text=f"처리 중... ({count}/{total_files})")

            progress.empty()

            if errors:
                for fname, err in errors:
                    st.error(f"{fname}: {err}")

            if generated_files:
                st.success(f"{len(generated_files)}개 파일 생성 완료!")

                # ZIP으로 묶어서 다운로드
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname, fpath in generated_files:
                        zf.write(fpath, fname)
                zip_buffer.seek(0)

                st.download_button(
                    label="📥 전체 파일 다운로드 (ZIP)",
                    data=zip_buffer,
                    file_name="한글파일_결과.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True,
                )

                # 개별 파일 다운로드
                with st.expander("개별 파일 다운로드"):
                    for fname, fpath in generated_files:
                        file_data = fpath.read_bytes()
                        st.download_button(
                            label=f"📄 {fname}",
                            data=file_data,
                            file_name=fname,
                            mime="application/octet-stream",
                            key=f"dl_{fname}",
                        )
else:
    st.divider()
    st.info("👆 위에서 템플릿과 엑셀 파일을 업로드하면 시작됩니다.")
