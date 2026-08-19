# -*- coding: utf-8 -*-
"""
시트 정리 프로그램 - 핵심 로직

엑셀 파일을 읽고/쓰고, 열을 매칭하고, 데이터를 복사하는 로직만 모아둔 파일.
tkinter(GUI)에 대한 의존성이 전혀 없어서, 화면 없이도 단독으로 테스트/재사용 가능.

GUI는 엑셀_정리_프로그램.py 에 있고, 이 파일의 함수들을 가져다 쓴다.
"""

import copy
import datetime
from collections import namedtuple

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string


# ------------------------------------------------------------------
# 기본값 (사용자가 "가져올 열" 칸을 비워두면 이 목록을 이름으로 매칭해서 사용)
# ------------------------------------------------------------------
DEFAULT_COLUMNS = [
    "No.", "계약년", "SYSTEM", "구분", "관리국", "SITE", "품질개선팀",
    "유니트명", "삼성SN", "PKG후장애", "고객 SN", "벤더SN",
    "입고일", "출고일", "수리년도",
]


# ------------------------------------------------------------------
# 유틸 함수
# ------------------------------------------------------------------

def normalize(name):
    if name is None:
        return ""
    return str(name).strip()


def build_header_map(ws, header_row):
    """{열이름: 열번호} 딕셔너리 생성. 원본/대상 어떤 시트에나 사용할 수 있다."""
    header_map = {}
    for col_idx in range(1, ws.max_column + 1):
        name = normalize(ws.cell(row=header_row, column=col_idx).value)
        if name:
            header_map[name] = col_idx
    return header_map


# ------------------------------------------------------------------
# 열 지정(column_spec) 파싱 & 해석
#
# 한 토큰(콤마로 구분)의 형태:  [출력이름=] [시트이름!] 원본식별자
#   "SITE"            -> 기준 시트에서 이름/번호/문자로 찾고, 출력 이름도 그대로 (하위호환)
#   "SITE=지역"        -> 기준 시트의 "지역" 열을 가져와 출력 이름은 "SITE"
#   "SITE=RMA!지역"    -> "RMA" 시트의 "지역" 열을 가져와 출력 이름은 "SITE"
#   "" (빈 토큰)        -> 이름도 데이터도 없는 스페이서 열
#   "비고="            -> 이름은 있지만 데이터는 없는 열 (수기 입력용)
# ------------------------------------------------------------------

RawToken = namedtuple("RawToken", ["output_name", "sheet_name", "ref"])
ColumnPlan = namedtuple("ColumnPlan", ["output_name", "sheet_name", "col_idx"])


def parse_column_spec(spec_text):
    """'가져올 열' 입력창의 텍스트를 문자열 수준에서만 파싱한다 (시트/헤더 조회 없음)."""
    tokens = []
    for raw in spec_text.split(","):
        t = raw.strip()
        if t == "":
            tokens.append(RawToken("", None, None))
            continue

        if "=" in t:
            left, right = t.split("=", 1)
            output_name = left.strip()
            right = right.strip()
        else:
            output_name = None
            right = t

        if right == "":
            tokens.append(RawToken(output_name or "", None, None))
            continue

        if "!" in right:
            sheet_part, ref_part = right.split("!", 1)
            sheet_name = sheet_part.strip() or None
            ref = ref_part.strip()
        else:
            sheet_name = None
            ref = right

        tokens.append(RawToken(output_name, sheet_name, ref))
    return tokens


def _resolve_single_ref(ref, header_map):
    """ref(이름/번호/열문자) 하나를 header_map에서 찾아 (열번호, 원본헤더이름)을 반환. 못 찾으면 (None, None)."""
    if ref in header_map:
        return header_map[ref], ref

    reverse_map = {v: k for k, v in header_map.items()}

    if ref.isdigit():
        col_idx = int(ref)
        return col_idx, reverse_map.get(col_idx, f"열{col_idx}")

    if ref.isalpha():
        try:
            col_idx = column_index_from_string(ref.upper())
            return col_idx, reverse_map.get(col_idx, f"열{col_idx}")
        except Exception:
            pass

    return None, None


def resolve_column_plan(wb, primary_sheet_name, header_row, spec_text, log, sheet_header_rows=None):
    """
    parse_column_spec 결과를 실제 워크북에 대해 해석해서 ColumnPlan 목록으로 변환한다.
    참조되는 시트마다 build_header_map을 지연 호출해서 캐싱한다.

    sheet_header_rows: {시트이름: 그 시트의 헤더 행 번호} (선택). 시트마다 헤더 행이 다를 수 있어서,
      여기 없는 시트는 기본값인 header_row(기준 시트의 헤더 행)를 그대로 쓴다.

    반환값: (plan, missing)
      plan: [ColumnPlan(output_name, sheet_name, col_idx), ...] (col_idx=None이면 스페이서 열)
      missing: 못 찾은 항목 설명 문자열 목록
    """
    sheet_header_rows = sheet_header_rows or {}
    header_maps = {}

    def get_header_map(sheet_name):
        if sheet_name not in header_maps:
            if sheet_name in wb.sheetnames:
                hr = sheet_header_rows.get(sheet_name, header_row)
                header_maps[sheet_name] = build_header_map(wb[sheet_name], hr)
            else:
                header_maps[sheet_name] = None
        return header_maps[sheet_name]

    plan = []
    missing = []

    if not spec_text or not spec_text.strip():
        # 기본 열 목록 사용 (기준 시트에서 이름 매칭)
        primary_map = get_header_map(primary_sheet_name) or {}
        for name in DEFAULT_COLUMNS:
            col_idx = primary_map.get(normalize(name))
            if col_idx:
                plan.append(ColumnPlan(name, primary_sheet_name, col_idx))
            else:
                missing.append(name)
        return plan, missing

    for tok in parse_column_spec(spec_text):
        if tok.ref is None:
            plan.append(ColumnPlan(tok.output_name, None, None))
            continue

        sheet_name = tok.sheet_name or primary_sheet_name
        hmap = get_header_map(sheet_name)
        if hmap is None:
            missing.append(f"{tok.ref} (시트 '{sheet_name}' 없음)")
            continue

        col_idx, resolved_name = _resolve_single_ref(tok.ref, hmap)
        if col_idx is None:
            missing.append(f"{tok.ref} ({sheet_name} 시트에서 못 찾음)" if tok.sheet_name else tok.ref)
            continue

        output_name = tok.output_name if tok.output_name is not None else resolved_name
        plan.append(ColumnPlan(output_name, sheet_name, col_idx))

    return plan, missing


# ------------------------------------------------------------------
# 여러 시트 결합 (키 값 기준 매칭)
# ------------------------------------------------------------------

def build_key_index(ws, header_row, key_col_idx, log, sheet_label):
    """시트를 훑어서 {정규화된 키 값: 행번호} 딕셔너리를 만든다. 중복 키는 첫 번째 행만 사용."""
    index = {}
    dup_count = 0
    for r in range(header_row + 1, ws.max_row + 1):
        v = normalize(ws.cell(row=r, column=key_col_idx).value)
        if not v:
            continue
        if v in index:
            dup_count += 1
            continue
        index[v] = r
    if dup_count:
        log(f"※ '{sheet_label}' 시트에 키 값이 중복된 행이 {dup_count}개 있어 첫 번째 행만 사용했습니다.")
    return index


def iter_matched_rows(wb, primary_sheet_name, header_row, plan, key_column_name, log, sheet_header_rows=None):
    """
    기준 시트를 행 단위로 순회하면서, plan의 각 항목에 대응하는 셀 위치를
    (시트이름, 행번호, 열번호) 튜플로 만들어 매 행마다 plan과 같은 길이의 리스트로 yield한다.
    (스페이서이거나 값을 못 찾으면 (None, None, None))

    sheet_header_rows: {시트이름: 그 시트의 헤더 행 번호} (선택). 시트마다 헤더 행이 다를 수 있어서,
      여기 없는 시트는 기본값인 header_row(기준 시트의 헤더 행)를 그대로 쓴다.

    다른 시트를 참조하는 항목이 있을 때:
      - key_column_name이 주어지면: 그 열 값이 같은 행끼리 짝을 맞춰서 가져온다 (순서가 달라도 안전)
      - key_column_name이 비어있으면: 그냥 "몇 번째 데이터 행인지"로 위치를 맞춰서 그대로 가져온다
        (기준 시트의 n번째 데이터 행 <-> 다른 시트의 n번째 데이터 행. 다른 시트도 같은 순서로
        정리되어 있다는 걸 사용자가 확인했을 때만 안전함)
    """
    sheet_header_rows = sheet_header_rows or {}
    ws_primary = wb[primary_sheet_name]
    primary_header_map = build_header_map(ws_primary, header_row)

    other_sheet_names = sorted({p.sheet_name for p in plan if p.sheet_name and p.sheet_name != primary_sheet_name})
    use_key = bool(key_column_name)

    key_col_primary_idx = None
    key_indexes = {}
    if use_key:
        key_col_primary_idx = primary_header_map.get(normalize(key_column_name))
        if key_col_primary_idx is None:
            raise ValueError(f"기준 열 '{key_column_name}'을(를) 기준 시트에서 찾을 수 없습니다.")

        for sheet_name in other_sheet_names:
            other_header_row = sheet_header_rows.get(sheet_name, header_row)
            other_header_map = build_header_map(wb[sheet_name], other_header_row)
            key_idx = other_header_map.get(normalize(key_column_name))
            if key_idx is None:
                raise ValueError(f"기준 열 '{key_column_name}'을(를) '{sheet_name}' 시트에서 찾을 수 없습니다.")
            key_indexes[sheet_name] = build_key_index(wb[sheet_name], other_header_row, key_idx, log, sheet_name)

    anchor = next(
        (p for p in plan if p.sheet_name == primary_sheet_name and p.col_idx is not None),
        None,
    )
    anchor_idx = anchor.col_idx if anchor else key_col_primary_idx
    if anchor_idx is None:
        raise ValueError("최소 1개 열은 기준 시트에서 가져오거나, 행 맞춤 기준 열을 지정해야 합니다.")

    unmatched_count = 0
    data_offset = 0  # 몇 번째 데이터 행인지(0부터) - 순서 기준 매칭에 사용
    for r in range(header_row + 1, ws_primary.max_row + 1):
        if ws_primary.cell(row=r, column=anchor_idx).value in (None, ""):
            continue

        key_val = normalize(ws_primary.cell(row=r, column=key_col_primary_idx).value) if key_col_primary_idx else None

        locations = []
        row_unmatched = False
        for p in plan:
            if p.col_idx is None:
                locations.append((None, None, None))
                continue
            if p.sheet_name == primary_sheet_name:
                locations.append((p.sheet_name, r, p.col_idx))
                continue
            if use_key:
                target_row = key_indexes.get(p.sheet_name, {}).get(key_val) if key_val else None
            else:
                other_header_row = sheet_header_rows.get(p.sheet_name, header_row)
                candidate_row = other_header_row + 1 + data_offset
                target_row = candidate_row if candidate_row <= wb[p.sheet_name].max_row else None
            if target_row is None:
                locations.append((None, None, None))
                row_unmatched = True
            else:
                locations.append((p.sheet_name, target_row, p.col_idx))
        if row_unmatched:
            unmatched_count += 1
        data_offset += 1
        yield locations

    if unmatched_count:
        if use_key:
            log(f"※ 다른 시트와 키 값이 맞지 않아 일부를 비워둔 행: {unmatched_count}개")
        else:
            log(f"※ 다른 시트에 대응하는 행이 부족해 일부를 비워둔 행: {unmatched_count}개")


def copy_cell(src_cell, dst_cell, copy_style, copy_formula):
    """src_cell의 내용을 dst_cell로 복사한다."""
    value = src_cell.value
    if copy_formula and isinstance(value, str) and value.startswith("="):
        dst_cell.value = value  # 수식 그대로 복사
    else:
        dst_cell.value = value

    if copy_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format


def read_raw_sheet(filepath, sheet_name, max_rows=300):
    """지정한 시트의 실제 내용을 그대로 읽어서 반환한다 (수식 셀은 계산된 값으로).
    반환값: (col_letters, rows) - col_letters는 ['A','B',...], rows는 각 행이 리스트인 목록
    각 행의 맨 앞에는 실제 엑셀 행 번호를 붙여준다.
    """
    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"'{sheet_name}' 시트를 찾을 수 없습니다.")
    ws = wb[sheet_name]

    max_col = ws.max_column or 1
    max_row = ws.max_row or 1
    col_letters = ["행"] + [get_column_letter(c) for c in range(1, max_col + 1)]

    rows = []
    for r_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=min(max_row, max_rows), max_col=max_col, values_only=True),
        start=1,
    ):
        rows.append([r_idx] + list(row))

    wb.close()
    return col_letters, rows, max_row


def get_sheet_names(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


def get_sheet_headers(filepath, sheet_name, header_row=1):
    """지정한 시트의 header_row에 있는 헤더 이름 목록을 순서대로 반환한다 (빈 칸 제외)."""
    wb = openpyxl.load_workbook(filepath, read_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f"'{sheet_name}' 시트를 찾을 수 없습니다.")
    ws = wb[sheet_name]
    names = []
    for col_idx in range(1, ws.max_column + 1):
        name = normalize(ws.cell(row=header_row, column=col_idx).value)
        if name:
            names.append(name)
    wb.close()
    return names


def make_unique_sheet_name(wb, base_name):
    if base_name not in wb.sheetnames:
        return base_name
    i = 1
    while f"{base_name}_{i}" in wb.sheetnames:
        i += 1
    return f"{base_name}_{i}"


def compute_result_rows(
    filepath,
    rma_sheet_name,
    rma_header_row,
    column_spec_text,
    target_mode,
    existing_sheet_name,
    log,
    key_column_name="",
    insert_ref_column="",
    insert_position="after",
    max_preview_rows=None,
    sheet_header_rows=None,
):
    """
    저장은 하지 않고, '실행'했을 때 나올 결과(헤더, 데이터 행)를 계산만 해서 돌려준다.
    미리보기와 실제 실행(run_conversion) 양쪽에서 공통으로 쓰는 계산 로직.

    반환값: (headers, rows, missing_cols, filled_names)
      - target_mode가 "existing"이면 headers는 기존 시트의 헤더를 기준으로 만들어짐
        (insert_ref_column/insert_position이 지정되면 그 위치에 새 열이 끼워진 모습으로 미리보기됨)
      - filled_names는 이번 실행에서 실제로 값이 채워지는 열 이름 목록
    """
    # 읽기 전용으로 열어서 엑셀이 열려있어도 항상 미리보기가 가능하게 함
    wb = openpyxl.load_workbook(filepath, read_only=False)

    if rma_sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(
            f"'{rma_sheet_name}' 시트를 찾을 수 없습니다. 현재 시트 목록: {wb.sheetnames}"
        )

    plan, missing_cols = resolve_column_plan(
        wb, rma_sheet_name, rma_header_row, column_spec_text, log, sheet_header_rows=sheet_header_rows
    )
    if not plan:
        wb.close()
        raise ValueError("가져올 열을 하나도 찾지 못했습니다. 입력값과 헤더 행 번호를 확인해주세요.")
    if missing_cols:
        log("※ 다음 열은 찾지 못해 제외됩니다: " + ", ".join(missing_cols))

    if target_mode == "existing" and existing_sheet_name in wb.sheetnames:
        ws_target = wb[existing_sheet_name]
        existing_header = build_header_map(ws_target, header_row=1)

        skipped = [p for p in plan if p.col_idx is None]
        if skipped:
            log(f"※ 빈 스페이서 열은 '기존 시트에 이어서 추가' 모드에서 지원하지 않아 제외됩니다 ({len(skipped)}개).")
        data_plan = [p for p in plan if p.col_idx is not None]

        headers = [None] * ws_target.max_column
        for name, idx in existing_header.items():
            headers[idx - 1] = name
        headers = [h if h else "" for h in headers]

        new_names = []
        for p in data_plan:
            if p.output_name not in existing_header and p.output_name not in new_names:
                new_names.append(p.output_name)
        if new_names:
            if insert_ref_column and insert_ref_column in existing_header:
                insert_at = existing_header[insert_ref_column] + (1 if insert_position == "after" else 0)
                headers = headers[:insert_at - 1] + new_names + headers[insert_at - 1:]
            else:
                headers = headers + new_names

        plan_by_name = {}
        for i, p in enumerate(data_plan):
            plan_by_name.setdefault(p.output_name, i)

        rows = []
        for locations in iter_matched_rows(
            wb, rma_sheet_name, rma_header_row, data_plan, key_column_name, log, sheet_header_rows=sheet_header_rows
        ):
            row_values = []
            for h in headers:
                plan_idx = plan_by_name.get(h)
                if plan_idx is None:
                    row_values.append("")
                    continue
                sheet, r, c = locations[plan_idx]
                row_values.append(wb[sheet].cell(row=r, column=c).value if c is not None else "")
            rows.append(row_values)
            if max_preview_rows and len(rows) >= max_preview_rows:
                break

        filled_names = [p.output_name for p in data_plan]

    else:
        data_plan = plan
        headers = [p.output_name for p in data_plan]
        rows = []
        for locations in iter_matched_rows(
            wb, rma_sheet_name, rma_header_row, data_plan, key_column_name, log, sheet_header_rows=sheet_header_rows
        ):
            row_values = [
                (wb[sheet].cell(row=r, column=c).value if c is not None else "")
                for sheet, r, c in locations
            ]
            rows.append(row_values)
            if max_preview_rows and len(rows) >= max_preview_rows:
                break
        filled_names = [p.output_name for p in data_plan]

    wb.close()
    return headers, rows, missing_cols, filled_names


# ------------------------------------------------------------------
# 핵심 처리
# ------------------------------------------------------------------

def run_conversion(
    filepath,
    rma_sheet_name,
    rma_header_row,
    column_spec_text,
    target_mode,          # "new" 또는 "existing"
    new_sheet_name,
    existing_sheet_name,
    copy_style,
    copy_formula,
    log,
    key_column_name="",
    insert_ref_column="",
    insert_position="after",
    sheet_header_rows=None,
):
    log(f"파일 여는 중: {filepath}")
    # 수식 문자열이 필요하므로 data_only=False (기본값) 로 연다.
    wb = openpyxl.load_workbook(filepath)

    if rma_sheet_name not in wb.sheetnames:
        raise ValueError(
            f"'{rma_sheet_name}' 시트를 찾을 수 없습니다. "
            f"현재 시트 목록: {wb.sheetnames}"
        )

    plan, missing_cols = resolve_column_plan(
        wb, rma_sheet_name, rma_header_row, column_spec_text, log, sheet_header_rows=sheet_header_rows
    )
    if not plan:
        raise ValueError("가져올 열을 하나도 찾지 못했습니다. 입력값과 헤더 행 번호를 확인해주세요.")
    if missing_cols:
        log("※ 다음 열은 찾지 못해 제외됩니다: " + ", ".join(missing_cols))

    # ---- 대상 시트 결정 ----
    if target_mode == "new":
        base_name = new_sheet_name.strip() or datetime.datetime.now().strftime("SYSTEM_%Y%m%d")
        sheet_name = make_unique_sheet_name(wb, base_name)
        ws_target = wb.create_sheet(title=sheet_name)
        log(f"새 시트 생성: {sheet_name}")

        data_plan = plan
        for i, p in enumerate(data_plan, start=1):
            ws_target.cell(row=1, column=i, value=p.output_name)
        target_col_positions = list(range(1, len(data_plan) + 1))
        start_row = 2

    else:  # existing
        if existing_sheet_name not in wb.sheetnames:
            raise ValueError(f"'{existing_sheet_name}' 시트를 찾을 수 없습니다.")
        ws_target = wb[existing_sheet_name]
        sheet_name = existing_sheet_name
        log(f"기존 시트 사용: {sheet_name}")

        skipped = [p for p in plan if p.col_idx is None]
        if skipped:
            log(f"※ 빈 스페이서 열은 '기존 시트에 이어서 추가' 모드에서 지원하지 않아 제외됩니다 ({len(skipped)}개).")
        data_plan = [p for p in plan if p.col_idx is not None]

        existing_header = build_header_map(ws_target, header_row=1)

        if not existing_header:
            # 기존 시트에 헤더가 아예 없으면 새로 씀
            for i, p in enumerate(data_plan, start=1):
                ws_target.cell(row=1, column=i, value=p.output_name)
            existing_header = {p.output_name: i for i, p in enumerate(data_plan, start=1)}
        else:
            new_names = []
            for p in data_plan:
                if p.output_name not in existing_header and p.output_name not in new_names:
                    new_names.append(p.output_name)

            if new_names:
                if insert_ref_column and insert_ref_column in existing_header:
                    # 이번에 새로 추가되는 열 전체를 기준 열 옆 한 지점에 묶어서 삽입
                    insert_at = existing_header[insert_ref_column] + (1 if insert_position == "after" else 0)
                    ws_target.insert_cols(insert_at, amount=len(new_names))
                    existing_header = {
                        name: (idx + len(new_names) if idx >= insert_at else idx)
                        for name, idx in existing_header.items()
                    }
                    for offset, name in enumerate(new_names):
                        col = insert_at + offset
                        ws_target.cell(row=1, column=col, value=name)
                        existing_header[name] = col
                    where = "뒤" if insert_position == "after" else "앞"
                    log(f"  '{insert_ref_column}' {where}에 새 열 {len(new_names)}개를 삽입했습니다: {', '.join(new_names)}")
                else:
                    # 지정 안 하면 기존처럼 맨 뒤에 추가
                    for name in new_names:
                        new_idx = ws_target.max_column + 1
                        ws_target.cell(row=1, column=new_idx, value=name)
                        existing_header[name] = new_idx
                        log(f"  '{name}' 열이 기존 시트에 없어 새로 추가했습니다.")

        target_col_positions = [existing_header[p.output_name] for p in data_plan]

        # 마지막 데이터 다음 행부터 이어서 추가.
        # ※ 시트 전체가 아니라, '이번에 실제로 채워 넣는 열들'만 봐서 마지막 행을 찾는다.
        last_row = 1
        for r in range(ws_target.max_row, 1, -1):
            if any(ws_target.cell(row=r, column=c).value not in (None, "") for c in target_col_positions):
                last_row = r
                break
        start_row = last_row + 1

    # ---- 데이터 복사 ----
    out_row = start_row
    copied_rows = 0
    for locations in iter_matched_rows(
        wb, rma_sheet_name, rma_header_row, data_plan, key_column_name, log, sheet_header_rows=sheet_header_rows
    ):
        for plan_idx, (sheet, r, c) in enumerate(locations):
            dst_cell = ws_target.cell(row=out_row, column=target_col_positions[plan_idx])
            if c is None:
                dst_cell.value = None
            else:
                copy_cell(wb[sheet].cell(row=r, column=c), dst_cell, copy_style, copy_formula)
        out_row += 1
        copied_rows += 1

    log(f"복사된 데이터 행: {copied_rows}개")
    log("=== 가져온 열 ===")
    for p in data_plan:
        src_desc = f"{p.sheet_name}!{get_column_letter(p.col_idx)}" if p.col_idx is not None else "(빈 열)"
        log(f"  {p.output_name}  <-  {src_desc}")

    wb.save(filepath)
    log("")
    log(f"저장 완료: {filepath}  (시트: {sheet_name})")
    return sheet_name
