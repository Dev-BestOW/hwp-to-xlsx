"""공통 유틸리티 함수 모듈."""

import re
from pathlib import Path


def generate_output_filename(
    data: dict[str, str],
    pattern: str,
    extension: str,
    index: int,
) -> str:
    """출력 파일명을 생성한다.

    Args:
        data: 행 데이터 딕셔너리
        pattern: 파일명 패턴 (예: "{이름}_output")
        extension: 파일 확장자 (예: ".hwpx")
        index: 행 인덱스 (0-based)

    Returns:
        생성된 파일명
    """
    try:
        filename = pattern.format(**data)
    except KeyError:
        # 패턴에 사용된 키가 데이터에 없으면 인덱스 기반 이름 사용
        filename = f"output_{index + 1}"

    # 파일명에 사용할 수 없는 문자 제거
    filename = sanitize_filename(filename)
    return filename + extension


def sanitize_filename(name: str) -> str:
    """파일명에 사용할 수 없는 문자를 제거한다."""
    return re.sub(r'[\\/*?:"<>|]', "_", name).strip()


def validate_template_extension(template_path: str | Path) -> str:
    """템플릿 파일의 확장자를 확인하고 반환한다.

    Returns:
        확장자 (소문자, 예: ".hwpx")

    Raises:
        ValueError: 지원하지 않는 확장자일 때
    """
    ext = Path(template_path).suffix.lower()
    supported = {".hwpx"}
    if ext not in supported:
        raise ValueError(
            f"지원하지 않는 템플릿 형식입니다: {ext}\n"
            f"지원 형식: {', '.join(supported)}"
        )
    return ext
