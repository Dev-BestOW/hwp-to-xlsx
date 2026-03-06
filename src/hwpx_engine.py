"""HWPX 템플릿의 플레이스홀더를 치환하는 엔진 모듈."""

import re
import shutil
from pathlib import Path

from hwpx import HwpxDocument


def replace_placeholders(
    template_path: str | Path,
    output_path: str | Path,
    data: dict[str, str],
) -> list[str]:
    """HWPX 템플릿의 플레이스홀더를 데이터로 치환하여 새 파일로 저장한다.

    Args:
        template_path: 템플릿 HWPX 파일 경로
        output_path: 출력 HWPX 파일 경로
        data: {플레이스홀더_이름: 치환값} 딕셔너리

    Returns:
        치환된 플레이스홀더 이름 목록

    Raises:
        FileNotFoundError: 템플릿 파일이 존재하지 않을 때
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 템플릿을 출력 경로에 복사한 뒤 수정
    shutil.copy2(template_path, output_path)

    doc = HwpxDocument.open(output_path)
    replaced = []

    for key, value in data.items():
        placeholder = "{{" + key + "}}"
        count = doc.replace_text_in_runs(placeholder, value)
        if count > 0:
            replaced.append(key)

    doc.save_to_path(output_path)
    return replaced


def find_placeholders_in_template(template_path: str | Path) -> list[str]:
    """템플릿 HWPX 파일에서 {{...}} 형식의 플레이스홀더를 찾아 반환한다.

    Args:
        template_path: 템플릿 HWPX 파일 경로

    Returns:
        플레이스홀더 이름 목록 (중복 제거)
    """
    template_path = Path(template_path)
    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")

    doc = HwpxDocument.open(template_path)
    text = doc.export_text()

    pattern = re.compile(r"\{\{(.+?)\}\}")
    matches = pattern.findall(text)
    return list(dict.fromkeys(matches))  # 순서 보존 중복 제거
