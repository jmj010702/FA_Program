# -*- coding: utf-8 -*-
"""
시트 정리 프로그램 - 핵심 로직

엑셀 파일을 읽고/쓰고, 열을 매칭하고, 데이터를 복사하는 로직만 모아둔 파일.
tkinter(GUI)에 대한 의존성이 전혀 없어서, 화면 없이도 단독으로 테스트/재사용 가능.

GUI는 엑셀_정리_프로그램.py 에 있고, 이 파일의 함수들을 가져다 쓴다.
"""

import copy
import datetime

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
    """{열이름: 열번호} 딕셔너리 생성 (RMA 시트용)"""
    header_map = {}
    for col_idx in range(1, ws.max_column + 1):
        name = normalize(ws.cell(row=header_row, column=col_idx).value)
        if name:
            header_map[name] = col_idx
    return header_map


def resolve_columns(spec_text, header_map):
    """
    '가져올 열' 입력창의 텍스트를 파싱해서
    [(output_name, rma_col_idx), ...] 리스트로 변환한다.

    입력 예시:
      "관리국,SITE,유니트명"   -> 이름으로 지정
      "1,3,5"                  -> 번호로 지정
      "A,C,E"                  -> 열 문자로 지정
      (비워두면) 기본 DEFAULT_COLUMNS 사용
    """
    result = []
    missing = []

    if not spec_text or not spec_text.strip():
        # 기본 열 목록 사용 (이름 매칭)
        for name in DEFAULT_COLUMNS:
            col_idx = header_map.get(normalize(name))
            if col_idx:
                result.append((name, col_idx))
            else:
                missing.append(name)
        return result, missing

    tokens = [t.strip() for t in spec_text.split(",") if t.strip()]
    reverse_map = {v: k for k, v in header_map.items()}  # 열번호 -> 이름

    for token in tokens:
        # 1) 이름으로 먼저 시도
        if token in header_map:
            result.append((token, header_map[token]))
            continue

        # 2) 숫자(열 번호)로 시도
        if token.isdigit():
            col_idx = int(token)
            out_name = reverse_map.get(col_idx, f"열{col_idx}")
            result.append((out_name, col_idx))
            continue

        # 3) 열 문자(A, B, C ...)로 시도
        if token.isalpha():
            try:
                col_idx = column_index_from_string(token.upper())
                out_name = reverse_map.get(col_idx, f"열{col_idx}")
                result.append((out_name, col_idx))
                continue
            except Exception:
                pass

        # 어느 것도 해당 안 되면 못 찾은 것으로 기록
        missing.append(token)

    return result, missing


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
    max_preview_rows=None,
):
    """
    저장은 하지 않고, '실행'했을 때 나올 결과(헤더, 데이터 행)를 계산만 해서 돌려준다.
    미리보기와 실제 실행(run_conversion) 양쪽에서 공통으로 쓰는 계산 로직.

    반환값: (headers, rows, missing_cols, filled_names)
      - target_mode가 "existing"이면 headers는 기존 시트의 헤더를 기준으로 만들어짐
      - filled_names는 이번 실행에서 실제로 값이 채워지는 열 이름 목록 (미리보기에서
        그 열로 자동 스크롤하는 데 사용)
    """
    # 읽기 전용으로 열어서 엑셀이 열려있어도 항상 미리보기가 가능하게 함
    wb = openpyxl.load_workbook(filepath, read_only=False)

    if rma_sheet_name not in wb.sheetnames:
        raise ValueError(
            f"'{rma_sheet_name}' 시트를 찾을 수 없습니다. 현재 시트 목록: {wb.sheetnames}"
        )
    ws_rma = wb[rma_sheet_name]

    header_map = build_header_map(ws_rma, rma_header_row)
    log(f"RMA 시트({rma_header_row}행)에서 찾은 열 이름 {len(header_map)}개: {list(header_map.keys())}")

    columns, missing_cols = resolve_columns(column_spec_text, header_map)
    if not columns:
        raise ValueError("가져올 열을 하나도 찾지 못했습니다. 입력값과 헤더 행 번호를 확인해주세요.")
    if missing_cols:
        log("※ 다음 열은 RMA 시트에서 찾지 못해 제외됩니다: " + ", ".join(missing_cols))

    # 대상(출력) 헤더 결정
    if target_mode == "existing" and existing_sheet_name in wb.sheetnames:
        ws_target = wb[existing_sheet_name]
        existing_header = {}
        for col_idx in range(1, ws_target.max_column + 1):
            name = normalize(ws_target.cell(row=1, column=col_idx).value)
            if name:
                existing_header[name] = col_idx
        if existing_header:
            # 기존 시트 헤더 순서 그대로 표시하고, 그 중 이번에 채워질 열만 데이터가 들어감
            headers = [None] * ws_target.max_column
            for name, idx in existing_header.items():
                headers[idx - 1] = name
            headers = [h if h else "" for h in headers]
        else:
            headers = []

        # 이번에 채워지는 열 중, 기존 시트에 아직 없는 열은 실제 실행 시와 동일하게
        # 맨 뒤에 새 열로 추가되므로 미리보기에도 똑같이 반영한다.
        existing_names = set(h for h in headers if h)
        for name, _ in columns:
            if name not in existing_names:
                headers.append(name)
                existing_names.add(name)
    else:
        headers = [name for name, _ in columns]

    # 각 출력 헤더 이름 -> RMA 열 인덱스 매핑 (없으면 None)
    name_to_rma_idx = {name: idx for name, idx in columns}

    data_start = rma_header_row + 1
    data_end = ws_rma.max_row
    first_col_idx = columns[0][1]

    rows = []
    for r in range(data_start, data_end + 1):
        if ws_rma.cell(row=r, column=first_col_idx).value in (None, ""):
            continue
        row_values = []
        for h in headers:
            rma_idx = name_to_rma_idx.get(h)
            if rma_idx:
                v = ws_rma.cell(row=r, column=rma_idx).value
            else:
                v = ""  # 이번에 채워지지 않는 기존 열, 또는 매칭 안 된 열
            row_values.append(v)
        rows.append(row_values)
        if max_preview_rows and len(rows) >= max_preview_rows:
            break

    wb.close()
    filled_names = [name for name, _ in columns]  # 이번에 실제로 값이 채워지는 열 이름들
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
):
    log(f"파일 여는 중: {filepath}")
    # 수식 문자열이 필요하므로 data_only=False (기본값) 로 연다.
    wb = openpyxl.load_workbook(filepath)

    if rma_sheet_name not in wb.sheetnames:
        raise ValueError(
            f"'{rma_sheet_name}' 시트를 찾을 수 없습니다. "
            f"현재 시트 목록: {wb.sheetnames}"
        )
    ws_rma = wb[rma_sheet_name]

    header_map = build_header_map(ws_rma, rma_header_row)
    log(f"RMA 시트({rma_header_row}행)에서 찾은 열 이름 {len(header_map)}개: {list(header_map.keys())}")

    columns, missing_cols = resolve_columns(column_spec_text, header_map)

    if not columns:
        raise ValueError("가져올 열을 하나도 찾지 못했습니다. 입력값과 헤더 행 번호를 확인해주세요.")

    if missing_cols:
        log("※ 다음 열은 RMA 시트에서 찾지 못해 제외됩니다: " + ", ".join(missing_cols))

    # ---- 대상 시트 결정 ----
    if target_mode == "new":
        base_name = new_sheet_name.strip() or datetime.datetime.now().strftime("SYSTEM_%Y%m%d")
        sheet_name = make_unique_sheet_name(wb, base_name)
        ws_target = wb.create_sheet(title=sheet_name)
        log(f"새 시트 생성: {sheet_name}")

        # 헤더 쓰기
        for i, (out_name, _) in enumerate(columns, start=1):
            ws_target.cell(row=1, column=i, value=out_name)
        target_col_positions = {out_name: i for i, (out_name, _) in enumerate(columns, start=1)}
        start_row = 2

    else:  # existing
        if existing_sheet_name not in wb.sheetnames:
            raise ValueError(f"'{existing_sheet_name}' 시트를 찾을 수 없습니다.")
        ws_target = wb[existing_sheet_name]
        sheet_name = existing_sheet_name
        log(f"기존 시트 사용: {sheet_name}")

        # 기존 시트의 1행을 헤더로 간주하고 이름 -> 열번호 매핑
        existing_header = {}
        for col_idx in range(1, ws_target.max_column + 1):
            name = normalize(ws_target.cell(row=1, column=col_idx).value)
            if name:
                existing_header[name] = col_idx

        if not existing_header:
            # 기존 시트에 헤더가 아예 없으면 새로 씀
            for i, (out_name, _) in enumerate(columns, start=1):
                ws_target.cell(row=1, column=i, value=out_name)
            target_col_positions = {out_name: i for i, (out_name, _) in enumerate(columns, start=1)}
        else:
            target_col_positions = {}
            for out_name, _ in columns:
                if out_name in existing_header:
                    target_col_positions[out_name] = existing_header[out_name]
                else:
                    # 기존 시트에 없는 열이면 맨 뒤에 새로 추가
                    new_idx = ws_target.max_column + 1
                    ws_target.cell(row=1, column=new_idx, value=out_name)
                    target_col_positions[out_name] = new_idx
                    log(f"  '{out_name}' 열이 기존 시트에 없어 새로 추가했습니다.")

        # 마지막 데이터 다음 행부터 이어서 추가.
        # ※ 시트 전체가 아니라, '이번에 실제로 채워 넣는 열들'만 봐서 마지막 행을 찾는다.
        #    (예: 이전엔 SITE 열만 채웠고 이번엔 No./계약년/SYSTEM 열을 추가하는 경우,
        #     SITE 열에 데이터가 있어도 무시하고 No./계약년/SYSTEM 열 기준으로 빈 자리부터 채움)
        target_cols_this_run = list(target_col_positions.values())
        last_row = 1
        for r in range(ws_target.max_row, 1, -1):
            if any(ws_target.cell(row=r, column=c).value not in (None, "") for c in target_cols_this_run):
                last_row = r
                break
        start_row = last_row + 1

    # ---- 데이터 복사 ----
    data_start = rma_header_row + 1
    data_end = ws_rma.max_row

    out_row = start_row
    copied_rows = 0
    for r in range(data_start, data_end + 1):
        # 빈 행은 건너뜀 (해당 행의 첫번째로 지정된 열이 비어있으면 스킵)
        first_col_idx = columns[0][1]
        if ws_rma.cell(row=r, column=first_col_idx).value in (None, ""):
            continue
        for out_name, rma_col_idx in columns:
            src_cell = ws_rma.cell(row=r, column=rma_col_idx)
            dst_col = target_col_positions[out_name]
            dst_cell = ws_target.cell(row=out_row, column=dst_col)
            copy_cell(src_cell, dst_cell, copy_style, copy_formula)
        out_row += 1
        copied_rows += 1

    log(f"복사된 데이터 행: {copied_rows}개")
    log("=== 가져온 열 ===")
    for out_name, rma_col_idx in columns:
        log(f"  {out_name}  <-  RMA {get_column_letter(rma_col_idx)}열")

    wb.save(filepath)
    log("")
    log(f"저장 완료: {filepath}  (시트: {sheet_name})")
    return sheet_name
