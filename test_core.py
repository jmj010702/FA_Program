# -*- coding: utf-8 -*-
"""
core.py 핵심 로직에 대한 최소 단위 테스트.

실행 방법:
    pytest

화면(tkinter)을 열지 않고, 함수 하나하나에 "이런 입력을 주면 이런 결과가 나와야 한다"를
코드로 확인한다. 기능을 추가하다가 기존 동작이 실수로 깨지면 여기서 바로 알 수 있다.
"""

import openpyxl

from core import (
    normalize,
    build_header_map,
    resolve_columns,
    copy_cell,
    make_unique_sheet_name,
)


# ------------------------------------------------------------------
# normalize
# ------------------------------------------------------------------

def test_normalize_공백_제거():
    assert normalize("  SITE  ") == "SITE"


def test_normalize_none은_빈문자열():
    assert normalize(None) == ""


def test_normalize_숫자도_문자열로():
    assert normalize(123) == "123"


# ------------------------------------------------------------------
# build_header_map
# ------------------------------------------------------------------

def test_build_header_map_헤더행에서_이름과_열번호_추출():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["No.", "SITE", "유니트명"])  # 1행

    header_map = build_header_map(ws, header_row=1)

    assert header_map == {"No.": 1, "SITE": 2, "유니트명": 3}


def test_build_header_map_빈_셀은_제외():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["No.", None, "유니트명"])

    header_map = build_header_map(ws, header_row=1)

    assert header_map == {"No.": 1, "유니트명": 3}


# ------------------------------------------------------------------
# resolve_columns
# ------------------------------------------------------------------

def test_resolve_columns_이름으로_지정():
    header_map = {"관리국": 1, "SITE": 2, "유니트명": 3}

    result, missing = resolve_columns("관리국,SITE", header_map)

    assert result == [("관리국", 1), ("SITE", 2)]
    assert missing == []


def test_resolve_columns_번호로_지정():
    header_map = {"관리국": 1, "SITE": 2}

    result, missing = resolve_columns("1,2", header_map)

    assert result == [("관리국", 1), ("SITE", 2)]
    assert missing == []


def test_resolve_columns_열문자로_지정():
    header_map = {"관리국": 1, "SITE": 2}

    result, missing = resolve_columns("A,B", header_map)

    assert result == [("관리국", 1), ("SITE", 2)]
    assert missing == []


def test_resolve_columns_찾을_수_없는_열은_missing으로():
    header_map = {"관리국": 1}

    result, missing = resolve_columns("관리국,존재안함", header_map)

    assert result == [("관리국", 1)]
    assert missing == ["존재안함"]


def test_resolve_columns_비어있으면_기본_열_목록_사용():
    header_map = {"No.": 1, "SITE": 2}  # DEFAULT_COLUMNS 중 일부만 존재

    result, missing = resolve_columns("", header_map)

    assert ("No.", 1) in result
    assert ("SITE", 2) in result
    assert "계약년" in missing  # header_map에 없는 기본 열은 missing으로


# ------------------------------------------------------------------
# copy_cell
# ------------------------------------------------------------------

def test_copy_cell_값만_복사():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "hello"
    src = ws["A1"]
    dst = ws["B1"]

    copy_cell(src, dst, copy_style=False, copy_formula=False)

    assert dst.value == "hello"


def test_copy_cell_수식복사_옵션_켜면_수식_그대로():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "=SUM(1,2)"
    src = ws["A1"]
    dst = ws["B1"]

    copy_cell(src, dst, copy_style=False, copy_formula=True)

    assert dst.value == "=SUM(1,2)"


def test_copy_cell_수식복사_옵션_끄면_수식_문자열_그대로_값처럼_복사():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "=SUM(1,2)"
    src = ws["A1"]
    dst = ws["B1"]

    copy_cell(src, dst, copy_style=False, copy_formula=False)

    # copy_formula=False 이면 수식 문자열을 그대로 값으로 취급해서 복사한다
    assert dst.value == "=SUM(1,2)"


# ------------------------------------------------------------------
# make_unique_sheet_name
# ------------------------------------------------------------------

def test_make_unique_sheet_name_중복없으면_그대로():
    wb = openpyxl.Workbook()

    assert make_unique_sheet_name(wb, "SYSTEM") == "SYSTEM"


def test_make_unique_sheet_name_중복되면_번호_붙임():
    wb = openpyxl.Workbook()
    wb.create_sheet("SYSTEM")
    wb.create_sheet("SYSTEM_1")

    assert make_unique_sheet_name(wb, "SYSTEM") == "SYSTEM_2"
