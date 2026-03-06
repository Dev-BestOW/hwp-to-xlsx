"""HWPX 템플릿의 플레이스홀더를 치환하는 엔진 모듈.

python-hwpx 대신 직접 ZIP/XML을 다루어 실제 한글에서 만든 HWPX 파일과의
호환성을 보장한다.
"""

import re
import shutil
import zipfile
from pathlib import Path
from lxml import etree


# HWPX 내부 XML에서 사용하는 네임스페이스
_NAMESPACES = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}

_HP_T = f"{{{_NAMESPACES['hp']}}}t"
_HP_RUN = f"{{{_NAMESPACES['hp']}}}run"
_HP_P = f"{{{_NAMESPACES['hp']}}}p"

_SECTION_PATTERN = re.compile(r"^Contents/section\d+\.xml$")
_PLACEHOLDER_PATTERN = re.compile(r"\{\{(.+?)\}\}")


def _get_section_paths(zf: zipfile.ZipFile) -> list[str]:
    """HWPX ZIP 내 섹션 XML 파일 경로를 반환한다."""
    return sorted(
        name for name in zf.namelist()
        if _SECTION_PATTERN.match(name)
    )


def _collect_run_text(run_elem: etree._Element) -> str:
    """run 요소 내 모든 <hp:t> 텍스트를 합쳐 반환한다."""
    parts = []
    for t_elem in run_elem.iter(_HP_T):
        if t_elem.text:
            parts.append(t_elem.text)
    return "".join(parts)


def _replace_in_section_xml(xml_bytes: bytes, data: dict[str, str]) -> tuple[bytes, list[str]]:
    """섹션 XML에서 플레이스홀더를 치환한다.

    단일 run 내 치환과, 여러 run에 걸쳐 분할된 플레이스홀더도 처리한다.

    Returns:
        (수정된 XML bytes, 치환된 키 목록)
    """
    parser = etree.XMLParser(recover=True)
    root = etree.fromstring(xml_bytes, parser=parser)
    replaced = set()

    # 1단계: 단일 <hp:t> 내 치환
    for t_elem in root.iter(_HP_T):
        if t_elem.text and "{{" in t_elem.text:
            original = t_elem.text
            for key, value in data.items():
                placeholder = "{{" + key + "}}"
                if placeholder in original:
                    original = original.replace(placeholder, value)
                    replaced.add(key)
            t_elem.text = original

    # 2단계: 여러 run에 걸쳐 분할된 플레이스홀더 처리
    for p_elem in root.iter(_HP_P):
        runs = list(p_elem.iter(_HP_RUN))
        if len(runs) < 2:
            continue

        # 전체 텍스트 조합 후 플레이스홀더 존재 확인
        combined = "".join(_collect_run_text(r) for r in runs)
        if "{{" not in combined:
            continue

        # 플레이스홀더가 남아있으면 run 단위 병합 치환 시도
        _fix_split_placeholders(runs, data, replaced)

    xml_out = etree.tostring(root, xml_declaration=True, encoding="UTF-8")
    return xml_out, list(replaced)


def _fix_split_placeholders(
    runs: list[etree._Element],
    data: dict[str, str],
    replaced: set[str],
) -> None:
    """여러 run에 걸쳐 분할된 {{...}} 플레이스홀더를 병합하여 치환한다."""
    # 각 run의 첫 번째 <hp:t> 요소와 텍스트를 수집
    run_info = []  # [(t_elem, text), ...]
    for run in runs:
        t_elem = run.find(f".//{_HP_T}")
        text = t_elem.text if t_elem is not None and t_elem.text else ""
        run_info.append((t_elem, text))

    combined = "".join(info[1] for info in run_info)
    if "{{" not in combined:
        return

    # 치환 수행
    new_combined = combined
    for key, value in data.items():
        placeholder = "{{" + key + "}}"
        if placeholder in new_combined:
            new_combined = new_combined.replace(placeholder, value)
            replaced.add(key)

    if new_combined == combined:
        return

    # 첫 번째 run에 전체 텍스트를 넣고 나머지 run의 텍스트를 비운다
    for i, (t_elem, _) in enumerate(run_info):
        if t_elem is None:
            continue
        if i == 0:
            t_elem.text = new_combined
        else:
            t_elem.text = ""


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
    """
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 템플릿 ZIP을 읽고, 섹션 XML만 수정하여 새 ZIP으로 저장
    all_replaced = set()

    with zipfile.ZipFile(template_path, "r") as zf_in:
        section_paths = _get_section_paths(zf_in)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                raw = zf_in.read(item.filename)

                if item.filename in section_paths:
                    modified_xml, replaced_keys = _replace_in_section_xml(raw, data)
                    all_replaced.update(replaced_keys)
                    zf_out.writestr(item, modified_xml)
                else:
                    zf_out.writestr(item, raw)

    return list(all_replaced)


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

    all_text = []
    with zipfile.ZipFile(template_path, "r") as zf:
        for section_path in _get_section_paths(zf):
            xml_bytes = zf.read(section_path)
            parser = etree.XMLParser(recover=True)
            root = etree.fromstring(xml_bytes, parser=parser)
            for t_elem in root.iter(_HP_T):
                if t_elem.text:
                    all_text.append(t_elem.text)

    combined = "".join(all_text)
    matches = _PLACEHOLDER_PATTERN.findall(combined)
    return list(dict.fromkeys(matches))
