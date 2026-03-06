"""Excel → HWPX 자동 입력 시스템 CLI."""

import argparse
import sys
from pathlib import Path

from src.excel_reader import read_excel
from src.hwpx_engine import replace_placeholders, find_placeholders_in_template
from src.utils import generate_output_filename, validate_template_extension


def main():
    parser = argparse.ArgumentParser(
        description="엑셀 데이터를 HWPX 템플릿에 자동으로 채워넣는 도구",
    )
    parser.add_argument(
        "--template", "-t",
        required=True,
        help="HWPX 템플릿 파일 경로",
    )
    parser.add_argument(
        "--data", "-d",
        required=True,
        help="엑셀 데이터 파일 경로 (.xlsx)",
    )
    parser.add_argument(
        "--output", "-o",
        default="./output",
        help="출력 디렉토리 경로 (기본값: ./output)",
    )
    parser.add_argument(
        "--filename-pattern", "-f",
        default="{이름}_output",
        help='출력 파일명 패턴 (기본값: "{이름}_output")',
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="실제 파일을 생성하지 않고 매칭 정보만 출력",
    )

    args = parser.parse_args()

    template_path = Path(args.template)
    data_path = Path(args.data)
    output_dir = Path(args.output)

    # 1. 확장자 확인
    try:
        ext = validate_template_extension(template_path)
    except ValueError as e:
        print(f"오류: {e}", file=sys.stderr)
        sys.exit(1)

    # 2. 엑셀 데이터 읽기
    try:
        rows = read_excel(data_path)
    except (FileNotFoundError, ValueError) as e:
        print(f"오류: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"엑셀 데이터: {len(rows)}개 행 로드")

    # 3. 템플릿 플레이스홀더 확인
    try:
        placeholders = find_placeholders_in_template(template_path)
    except FileNotFoundError as e:
        print(f"오류: {e}", file=sys.stderr)
        sys.exit(1)

    if placeholders:
        print(f"템플릿 플레이스홀더: {', '.join(placeholders)}")
    else:
        print("경고: 템플릿에서 플레이스홀더를 찾을 수 없습니다.", file=sys.stderr)

    # 엑셀 헤더와 플레이스홀더 매칭 확인
    excel_headers = set(rows[0].keys()) if rows else set()
    matched = set(placeholders) & excel_headers
    unmatched = set(placeholders) - excel_headers
    if unmatched:
        print(f"경고: 매칭되지 않는 플레이스홀더: {', '.join(unmatched)}", file=sys.stderr)

    if args.dry_run:
        print(f"\n[Dry Run] 매칭된 필드: {', '.join(matched)}")
        print(f"[Dry Run] 생성될 파일 수: {len(rows)}")
        for i, row in enumerate(rows):
            filename = generate_output_filename(row, args.filename_pattern, ext, i)
            print(f"  - {filename}")
        sys.exit(0)

    # 4. 출력 디렉토리 생성
    output_dir.mkdir(parents=True, exist_ok=True)

    # 5. 각 행마다 HWPX 파일 생성
    success_count = 0
    for i, row in enumerate(rows):
        filename = generate_output_filename(row, args.filename_pattern, ext, i)
        output_path = output_dir / filename

        try:
            replaced = replace_placeholders(template_path, output_path, row)
            success_count += 1
            print(f"  [{i + 1}/{len(rows)}] {filename} (치환: {len(replaced)}개 필드)")
        except Exception as e:
            print(f"  [{i + 1}/{len(rows)}] 오류 - {filename}: {e}", file=sys.stderr)

    print(f"\n완료: {success_count}/{len(rows)}개 파일 생성 → {output_dir}/")


if __name__ == "__main__":
    main()
