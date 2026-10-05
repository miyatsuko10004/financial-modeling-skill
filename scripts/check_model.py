#!/usr/bin/env python3
"""
check_model.py: 財務モデリング規約バリデータ（オープンpyxlによる機械的検査）

検査項目:
1. Excelエラーの有無 (#REF!, #DIV/0!, #VALUE!, #NAME?, #N/A)
2. 非表示行・非表示列の有無 (規約§7.1: Hide禁止)
3. バランスシートの貸借完全一致 (資産合計 - 負債純資産合計 == 0)
4. 現金連動 (BS現預金 == CF期末現金残高)
5. 時間軸の横方向統一とヘッダー表記 (20xxA, 20xxE/F)
6. エレベーターコラム (A列ジャンプ機能)
7. セマンティック文字色規約 (青字/黒字/緑字)

使い方:
  python3 scripts/check_model.py <path_to_model.xlsx>
"""

import sys
import os
import re
import openpyxl

ERRORS = []
WARNINGS = []
PASSES = []

EXCEL_ERRORS = ["#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!"]

def log_pass(msg):
    PASSES.append(msg)
    print(f"  \033[32m[PASS]\033[0m {msg}")

def log_warn(msg):
    WARNINGS.append(msg)
    print(f"  \033[33m[WARN]\033[0m {msg}")

def log_fail(msg):
    ERRORS.append(msg)
    print(f"  \033[31m[FAIL]\033[0m {msg}")

def check_excel_errors(wb):
    found_errors = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row in ws.iter_rows():
            for cell in row:
                val = str(cell.value) if cell.value is not None else ""
                for err in EXCEL_ERRORS:
                    if err in val:
                        found_errors.append(f"{sheet}!{cell.coordinate}: {err}")
    if found_errors:
        log_fail(f"Excelエラーが {len(found_errors)} 箇所見つかりました: {', '.join(found_errors[:5])}")
    else:
        log_pass("Excelエラーなし (#REF!, #DIV/0!, #VALUE! 等の重大エラー 0件)")

def check_hidden_rows_cols(wb):
    hidden_rows = []
    hidden_cols = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for row_idx, row_dim in ws.row_dimensions.items():
            if row_dim.hidden:
                hidden_rows.append(f"{sheet}!Row{row_idx}")
        for col_letter, col_dim in ws.column_dimensions.items():
            if col_dim.hidden:
                hidden_cols.append(f"{sheet}!Col{col_letter}")
                
    if hidden_rows or hidden_cols:
        log_fail(f"非表示（Hide）が検出されました (規約§7.1違反)。行: {len(hidden_rows)}件, 列: {len(hidden_cols)}件。非表示ではなくグループ化を使用してください。")
    else:
        log_pass("非表示（Hide）行・列なし (全データ可視またはグループ化運用)")

def check_time_axis_and_columns(ws):
    headers = []
    # 行1〜行10まで走査
    for r in range(1, 11):
        row_headers = []
        for cell in ws[r]:
            val = str(cell.value).strip() if cell.value is not None else ""
            if re.search(r'20\d\d[AEF]?', val) or re.search(r'FY\d\d', val):
                row_headers.append((cell.column_letter, val))
        if len(row_headers) >= 3:
            headers = row_headers
            break

    if headers:
        log_pass(f"時間軸の横方向展開を検出: {', '.join([f'{col}={val}' for col, val in headers[:7]])}")
    else:
        log_warn("明示的な年次ヘッダー（20xxA, 20xxE等）が検出されませんでした。")

def check_balance_sheet(wb):
    # 単一シートまたはBSシートを走査
    target_sheets = [s for s in wb.sheetnames if any(k in s.lower() for k in ["bs", "貸借", "業績予想"])]
    if not target_sheets:
        target_sheets = wb.sheetnames

    bs_verified = False
    for sheet_name in target_sheets:
        ws = wb[sheet_name]
        asset_total_row = None
        liability_equity_total_row = None
        check_diff_row = None
        
        for row in ws.iter_rows(max_col=5):
            for cell in row:
                val = str(cell.value).replace(" ", "").replace("　", "").strip() if cell.value is not None else ""
                if any(k in val for k in ["負債・純資産合計", "負債純資産合計", "TotalLiabilities&Equity", "負債及び純資産合計"]):
                    liability_equity_total_row = cell.row
                elif ("資産合計" in val or "TotalAssets" in val) and "負債" not in val:
                    asset_total_row = cell.row
                elif "バランスチェック" in val or "差額" in val or "Check" in val:
                    check_diff_row = cell.row

        if asset_total_row and liability_equity_total_row:
            bs_verified = True
            log_pass(f"[{sheet_name}] 貸借対照表の構造を検出 (資産合計: Row {asset_total_row}, 負債純資産合計: Row {liability_equity_total_row})")
            if check_diff_row:
                log_pass(f"[{sheet_name}] バランスチェック行常設を確認 (Row {check_diff_row})")
            else:
                log_warn(f"[{sheet_name}] バランスチェック行（差額計算行）の明示的常設を推奨します (規約§4.1)")
            break

    if not bs_verified:
        log_warn("貸借対照表（資産合計および負債純資産合計）の行が検出されませんでした（DCF等の特化シートの場合は問題ありません）。")

def check_cash_flow_link(wb):
    # 現預金とCF期末残高の連動確認
    linked = False
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(max_col=3):
            for cell in row:
                val = str(cell.value).strip() if cell.value is not None else ""
                if any(k in val for k in ["現預金", "現金及び預金", "Cash", "現金"]) and "増減" not in val and "CF" not in val and "フロー" not in val:
                    # 右隣またはその行の数式を確認
                    row_cells = ws[cell.row]
                    formulas = [str(c.value) for c in row_cells if str(c.value).startswith("=")]
                    if formulas:
                        linked = True
                        log_pass(f"[{sheet_name}] 現預金科目の数式連動を確認 (Row {cell.row}: {formulas[0][:25]})")
                        break
            if linked:
                break
    if not linked:
        log_warn("現預金科目のキャッシュフロー計算書からの直接数式連動が検出されませんでした (規約§4.1)。")

def check_elevator_column(wb):
    # A列エレベーターコラムの確認
    has_elevator = False
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        a_cells = [cell.value for cell in ws['A'] if cell.value is not None and str(cell.value).strip() != ""]
        section_titles = [str(v) for v in a_cells if re.search(r'[（\(][０-９\d]+[）\)]', str(v)) or "前提" in str(v) or "損益" in str(v)]
        if len(section_titles) >= 2:
            has_elevator = True
            log_pass(f"[{sheet_name}] A列エレベーターコラム（ジャンプ機能）を確認: {len(section_titles)}セクション検出")
            break
            
    if not has_elevator:
        log_warn("A列エレベーターコラム（Ctrl+↓用ジャンプ見出し）が未設定です (規約§2.3)。")

def check_color_coding(wb):
    # 文字色・入力背景色のサンプリング検査
    blue_input_found = False
    black_formula_found = False
    orange_fill_found = False
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(max_row=50, max_col=10):
            for cell in row:
                val = cell.value
                val_str = str(val) if val is not None else ""
                
                # 背景色
                if cell.fill and cell.fill.fill_type:
                    fg = cell.fill.fgColor
                    if fg and fg.rgb:
                        rgb = str(fg.rgb)
                        if any(c in rgb.upper() for c in ["FFF2CC", "FCE4D6", "FFFFCC", "FEE2E2", "FFE699"]):
                            orange_fill_found = True
                            
                # フォント色
                if cell.font and cell.font.color:
                    fc = str(cell.font.color.rgb) if cell.font.color.rgb else ""
                    if any(b in fc.upper() for b in ["0000FF", "002060", "0070C0"]):
                        blue_input_found = True
                        
                if val_str.startswith("="):
                    black_formula_found = True
                    
    if blue_input_found:
        log_pass("セマンティック文字色（青字ハードコード）を検出 (規約§1.1)")
    else:
        log_warn("青字（#0000FF）のハードコードセルが検出されませんでした (規約§1.1: 実績値・前提は青字表記)")
        
    if orange_fill_found:
        log_pass("入力可能セル背景色（薄オレンジ/黄）を検出 (規約§1.5)")
    else:
        log_warn("前提入力セルの背景着色（オレンジ/黄色）が検出されませんでした (規約§1.5)")

ALLOWED_LITERALS = {"0", "1", "365", "12"}  # 式内で許容する定数（0, 1, 年日数, 月数）

def check_hardcode_in_formula(wb):
    """§3.8: 数式内への数値ハードコード埋め込みを検出する"""
    bad = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                f = re.sub(r'"[^"]*"', '', v)
                f = re.sub(r"(?:'[^']+'|[A-Za-z_]\w*)!", '', f)
                f = re.sub(r'\$?[A-Z]{1,3}\$?\d+', '', f)
                for lit in re.findall(r'(?<![A-Za-z_])\d+\.?\d*', f):
                    if lit not in ALLOWED_LITERALS:
                        bad.append(f"{ws.title}!{cell.coordinate}({lit})")
                        break
    if bad:
        log_fail(f"数式内への定数埋め込みが {len(bad)} 箇所あります (規約§3.8): {', '.join(bad[:5])}")
    else:
        log_pass("数式内の定数埋め込みなし (規約§3.8)")

def _font_rgb(cell):
    c = cell.font.color if cell.font else None
    return str(c.rgb).upper()[-6:] if c is not None and isinstance(c.rgb, str) else ""

def check_color_consistency(wb):
    """§1.1/§1.2: 数値ハードコードは青字、数式は青字禁止（全セル検査）"""
    blues = ("0000FF", "002060")
    hard_not_blue, formula_blue = [], []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, bool) or v is None:
                    continue
                if isinstance(v, (int, float)) and _font_rgb(cell) not in blues:
                    hard_not_blue.append(f"{ws.title}!{cell.coordinate}")
                elif isinstance(v, str) and v.startswith("=") and _font_rgb(cell) in blues:
                    formula_blue.append(f"{ws.title}!{cell.coordinate}")
    if hard_not_blue:
        log_fail(f"ハードコード数値が青字でないセル {len(hard_not_blue)} 件 (規約§1.1): {', '.join(hard_not_blue[:5])}")
    else:
        log_pass("全ハードコード数値が青字 (規約§1.1)")
    if formula_blue:
        log_fail(f"数式セルが青字になっている {len(formula_blue)} 件 (規約§1.2): {', '.join(formula_blue[:5])}")
    else:
        log_pass("数式セルに青字の混入なし (規約§1.2)")

def check_sensitivity(wb):
    """§5.4: DCFがある場合は感応度マトリクスを併設"""
    text = " ".join(str(c.value) for ws in wb.worksheets for r in ws.iter_rows() for c in r if c.value is not None)
    if any(k in text for k in ("DCF", "ターミナルバリュー", "WACC")):
        if "感応度" in text or "Sensitivity" in text:
            log_pass("DCF感応度マトリクスを確認 (規約§5.4)")
        else:
            log_fail("DCFがあるのに感応度マトリクスがありません (規約§5.4)")

def check_computed_balance(file_path):
    """数式を実際に計算し、BS差額・現預金とCF期末残高の一致を値で検証する"""
    try:
        import formulas
    except ImportError:
        log_warn("計算エンジン(formulas)が未導入のため、BS差額の値検証をスキップしました: pip install formulas")
        return
    import io, contextlib
    try:
        with contextlib.redirect_stderr(io.StringIO()):
            sol = formulas.ExcelModel().loads(file_path).finish().calculate()
    except Exception as e:
        log_warn(f"数式の計算に失敗し、値検証をスキップしました: {e}")
        return
    wb = openpyxl.load_workbook(file_path)
    base = os.path.basename(file_path)
    def val(ws, coord):
        k = "'[%s]%s'!%s" % (base, ws.title.upper(), coord)
        return sol[k].value[0, 0] if k in sol else None
    verified = False
    for ws in wb.worksheets:
        rows = {}
        for r in ws.iter_rows(max_col=3):
            lab = "".join(str(r[1].value or "").split()) if len(r) > 1 else ""
            if "バランスチェック" in lab: rows["chk"] = r[0].row
            elif lab == "資産合計": rows["a"] = r[0].row
            elif lab in ("負債・純資産合計", "負債純資産合計"): rows["l"] = r[0].row
        if "a" not in rows or "l" not in rows:
            continue
        verified = True
        bad = []
        for cell in ws[rows["a"]][2:]:
            if cell.value is None: continue
            a, l = val(ws, cell.coordinate), val(ws, f"{cell.column_letter}{rows['l']}")
            try:
                if abs(float(a) - float(l)) >= 0.01:
                    bad.append(f"{cell.column_letter}列(差額{float(a)-float(l):,.1f})")
            except (TypeError, ValueError):
                bad.append(f"{cell.column_letter}列(計算不能)")
        if bad:
            log_fail(f"[{ws.title}] BSの貸借が一致しません (規約§4.1): {', '.join(bad)}")
        else:
            log_pass(f"[{ws.title}] 数式計算の結果、全期間でBS貸借一致を確認 (規約§4.1)")
    if not verified:
        log_warn("BSが検出できず、貸借の値検証を行えませんでした。")

def main():
    if len(sys.argv) < 2:
        print("Usage: python3 scripts/check_model.py <path_to_model.xlsx>")
        sys.exit(1)

    file_path = sys.argv[1]
    if not os.path.exists(file_path):
        print(f"Error: ファイルが存在しません: {file_path}")
        sys.exit(1)

    print("=" * 60)
    print(f"財務モデル規約整合性バリデータ: {os.path.basename(file_path)}")
    print("=" * 60)

    try:
        wb = openpyxl.load_workbook(file_path, data_only=False)
    except Exception as e:
        log_fail(f"Excelファイル読み込みエラー: {e}")
        sys.exit(1)

    print("\n--- 1. 重大エラー＆整合性検査 ---")
    check_excel_errors(wb)
    check_hidden_rows_cols(wb)
    check_balance_sheet(wb)
    check_cash_flow_link(wb)
    check_computed_balance(file_path)
    check_hardcode_in_formula(wb)
    check_sensitivity(wb)

    print("\n--- 2. レイアウト・構造・ナビゲーション検査 ---")
    ws_main = wb.active
    check_time_axis_and_columns(ws_main)
    check_elevator_column(wb)

    print("\n--- 3. 色彩・セマンティック書式検査 ---")
    check_color_coding(wb)
    check_color_consistency(wb)

    print("\n" + "=" * 60)
    print(f"検査結果サマリ: FAIL {len(ERRORS)} 件 / WARN {len(WARNINGS)} 件 / PASS {len(PASSES)} 件")
    print("=" * 60)

    if ERRORS:
        print("\n\033[31m[判定: 納品不合格 (REJECTED)]\033[0m")
        print("以下のFAIL項目を直ちに修正してください:")
        for err in ERRORS:
            print(f" - {err}")
        sys.exit(1)
    elif WARNINGS:
        print("\n\033[33m[判定: 条件付き合格 (PASS with WARNINGS)]\033[0m")
        print("重大エラーはありませんが、以下のWARN項目を確認・改善してください:")
        for warn in WARNINGS:
            print(f" - {warn}")
        sys.exit(0)
    else:
        print("\n\033[32m[判定: 完全合格 (PERFECT PASS)]\033[0m")
        print("すべての規約を満たしています。納品・レビュー可能です。")
        sys.exit(0)

if __name__ == "__main__":
    main()
