# -*- coding: utf-8 -*-
"""Excel 读写：本地模板解析 + 通过 Excel COM 写入 SharePoint 在线工作簿。"""

import re
import traceback
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import openpyxl

try:
    import pythoncom
    import win32com.client
    from pywintypes import com_error
except ImportError:  # 非 Windows 环境
    pythoncom = None
    win32com = None
    com_error = Exception

XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_PASTE_FORMATS = -4122
XL_SHIFT_DOWN = -4162

REQUIRED_FIELDS = ("ticket", "description")
FIELD_ORDER = (
    "ticket",
    "description",
    "author",
    "repo",
    "branch",
    "date",
    "commits",
    "hours",
    "progress",
    "remark",
)


class NullLogger:
    def info(self, message):
        pass

    def warn(self, message):
        pass

    def error(self, message):
        pass


@dataclass
class WritePlan:
    sheet_name: str
    header_row: int
    start_row: int
    end_row: int
    columns: dict
    footer_row: int = None
    last_data_row: int = None

    def to_dict(self):
        return {
            "sheet_name": self.sheet_name,
            "header_row": self.header_row,
            "start_row": self.start_row,
            "end_row": self.end_row,
            "columns": self.columns,
            "footer_row": self.footer_row,
            "last_data_row": self.last_data_row,
        }


@dataclass
class WriteResult:
    plan: WritePlan
    snapshot_path: str = ""
    saved: bool = False
    read_only: bool = False
    escaped: int = 0
    verified: bool = False
    warnings: list = field(default_factory=list)

    def to_dict(self):
        return {
            "plan": self.plan.to_dict(),
            "snapshot_path": self.snapshot_path,
            "saved": self.saved,
            "read_only": self.read_only,
            "escaped": self.escaped,
            "verified": self.verified,
            "warnings": self.warnings,
        }


def normalize_header(text):
    if text is None:
        return ""
    value = str(text).replace("\u3000", " ").replace("\r", " ").replace("\n", " ")
    value = re.sub(r"\s+", "", value).lower()
    value = value.replace("（", "(").replace("）", ")").replace("：", ":")
    return value


def normalize_sharepoint_url(url):
    return re.sub(r"/:[a-z]:/r/", "/", str(url or "").strip())


def workbook_name_from_url(url):
    try:
        parsed = urlparse(str(url or ""))
        values = parse_qs(parsed.query).get("file")
        if values:
            return Path(unquote(values[0])).name
        if parsed.path:
            return Path(unquote(parsed.path)).name
    except (ValueError, TypeError):
        pass
    return "sharepoint.xlsx"


def weekly_sheet_name(prefix, range_start):
    if isinstance(range_start, str):
        try:
            start = datetime.strptime(range_start[:10], "%Y-%m-%d").date()
        except ValueError:
            start = datetime.now().date()
    elif isinstance(range_start, datetime):
        start = range_start.date()
    elif isinstance(range_start, date):
        start = range_start
    else:
        start = datetime.now().date()
    prefix = (prefix or "周报汇总").strip()
    separator = "" if prefix.endswith((" ", "-", "_")) else " "
    base = "{}{}{}.{}".format(prefix, separator, start.month, start.day)
    base = re.sub(r'[:\\/?*\[\]]', '_', base)
    return base[:31] or "周报汇总 {}.{}".format(start.month, start.day)


def _auto_weekly_template(workbook):
    candidates = []
    for sheet in workbook.Worksheets:
        name = str(sheet.Name)
        if "周报汇总" in name or "汇总" in name:
            candidates.append(name)
    return candidates[-1] if candidates else ""


def _alias_index(aliases):
    return {
        field: {normalize_header(name) for name in names if str(name).strip()}
        for field, names in aliases.items()
    }


def _match_header_row(values, alias_index):
    columns = {}
    for index, raw in enumerate(values, 1):
        text = normalize_header(raw)
        if not text:
            continue
        for field, names in alias_index.items():
            if field not in columns and text in names:
                columns[field] = index
    return columns


def _row_matches(columns):
    if len(columns) < 2:
        return False
    return any(field in columns for field in REQUIRED_FIELDS)


def _pick_best_header(rows, alias_index, forced_row=0):
    best = None
    for row_number, values in rows:
        if forced_row and row_number != forced_row:
            continue
        columns = _match_header_row(values, alias_index)
        if forced_row:
            return row_number, columns
        if _row_matches(columns):
            score = len(columns)
            if best is None or score > best[1]:
                best = (row_number, score, columns)
    if best is None:
        return None, {}
    return best[0], best[2]


def _text_preview(value, limit=120):
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ").strip()
    return text[:limit]


def inspect_local(path, aliases, forced_header_row=0, preview_rows=12, preview_cols=30):
    workbook = openpyxl.load_workbook(path, data_only=True, read_only=False)
    alias_index = _alias_index(aliases)
    sheets = []
    detected = None
    try:
        for sheet in workbook.worksheets:
            max_row = min(sheet.max_row or 1, 5000)
            max_col = min(sheet.max_column or 1, 200)
            scan_last_row = min(max_row, 40)
            scan_last_col = min(max_col, 80)
            scan = [
                (
                    row_number,
                    [
                        sheet.cell(row=row_number, column=column).value
                        for column in range(1, scan_last_col + 1)
                    ],
                )
                for row_number in range(1, scan_last_row + 1)
            ]
            header_row, columns = _pick_best_header(scan, alias_index, forced_header_row)
            preview = []
            for row in sheet.iter_rows(
                min_row=1,
                max_row=min(max_row, preview_rows),
                max_col=min(max_col, preview_cols),
                values_only=True,
            ):
                preview.append([_text_preview(value, 80) for value in row])
            item = {
                "name": sheet.title,
                "max_row": sheet.max_row,
                "max_column": sheet.max_column,
                "header_row": header_row,
                "columns": columns,
                "preview": preview,
            }
            sheets.append(item)
            if header_row and (
                detected is None or len(columns) > len(detected["columns"])
            ):
                detected = item
    finally:
        workbook.close()
    return {"path": str(path), "sheets": sheets, "detected": detected}


def _com_available():
    return pythoncom is not None and win32com is not None


def _require_com():
    if not _com_available():
        raise RuntimeError("当前环境不支持 Excel COM（需要 Windows + Excel + pywin32）")


def _start_excel():
    _require_com()
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.AskToUpdateLinks = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        return app
    except Exception:
        pythoncom.CoUninitialize()
        raise


def _stop_excel(app):
    try:
        app.Quit()
    except Exception:
        pass
    finally:
        pythoncom.CoUninitialize()


def _com_value(value):
    if hasattr(value, "__com_get_value__"):
        return value.__com_get_value__()
    return value


def _attach_via_window():
    try:
        import ctypes
        import uuid
        import win32gui
        import comtypes
        import comtypes.client
        from ctypes import wintypes
    except ImportError:
        raise RuntimeError("需要安装 comtypes 才能使用窗口附着模式")

    class GUID(ctypes.Structure):
        _fields_ = [
            ("Data1", wintypes.DWORD),
            ("Data2", wintypes.WORD),
            ("Data3", wintypes.WORD),
            ("Data4", ctypes.c_ubyte * 8),
        ]

    def make_guid(value):
        parsed = uuid.UUID(value)
        return GUID(
            parsed.time_low,
            parsed.time_mid,
            parsed.time_hi_version,
            (ctypes.c_ubyte * 8)(*parsed.bytes[8:]),
        )

    iid_dispatch = make_guid("{00020400-0000-0000-C000-000000000046}")
    objid_nativeom = -16
    oleacc = ctypes.windll.oleacc
    oleacc.AccessibleObjectFromWindow.argtypes = [
        wintypes.HWND,
        wintypes.DWORD,
        ctypes.POINTER(GUID),
        ctypes.POINTER(ctypes.c_void_p),
    ]
    oleacc.AccessibleObjectFromWindow.restype = ctypes.c_long

    main_windows = []

    def collect_main(hwnd, _):
        if win32gui.GetClassName(hwnd) == "XLMAIN":
            main_windows.append(hwnd)

    win32gui.EnumWindows(collect_main, None)
    if not main_windows:
        raise RuntimeError("未找到 Excel 主窗口")

    last_error = None
    for main in main_windows:
        grids = []

        def collect_grid(hwnd, _):
            if win32gui.GetClassName(hwnd) == "EXCEL7":
                grids.append(hwnd)

        win32gui.EnumChildWindows(main, collect_grid, None)
        for grid in grids:
            pointer = ctypes.c_void_p()
            result = oleacc.AccessibleObjectFromWindow(
                grid,
                objid_nativeom,
                ctypes.byref(iid_dispatch),
                ctypes.byref(pointer),
            )
            if result != 0 or not pointer.value:
                continue
            try:
                unknown = ctypes.cast(pointer, ctypes.POINTER(comtypes.IUnknown))
                window = comtypes.client.GetBestInterface(unknown)
                app = window.Application
                count = int(_com_value(app.Workbooks.Count))
                for index in range(1, count + 1):
                    workbook = app.Workbooks(index)
                    full_name = str(_com_value(workbook.FullName))
                    try:
                        py_workbook = win32com.client.GetObject(full_name)
                        return py_workbook.Application
                    except Exception as error:
                        last_error = error
            except Exception as error:
                last_error = error
    raise RuntimeError(
        "未能附着到 Excel 工作簿: {}".format(last_error or "未知错误")
    )


def _attach_excel(target_path=""):
    _require_com()
    pythoncom.CoInitialize()
    errors = []
    for label, getter in (
        (
            "GetActiveObject",
            lambda: win32com.client.GetActiveObject("Excel.Application"),
        ),
    ):
        try:
            app = getter()
            if int(app.Workbooks.Count) > 0:
                return app
            errors.append("{}: 没有打开的工作簿".format(label))
        except Exception as error:
            errors.append("{}: {}".format(label, error))
    if target_path:
        try:
            workbook = win32com.client.GetObject(str(target_path))
            return workbook.Application
        except Exception as error:
            errors.append("GetObject(path): {}".format(error))
    try:
        return _attach_via_window()
    except Exception as error:
        errors.append("窗口附着: {}".format(error))
    pythoncom.CoUninitialize()
    detail = "；".join(errors[-2:]) if errors else ""
    raise RuntimeError(
        "未检测到可自动化的 Excel。请先在 Excel 中正常打开周报"
        "（不要用 SharePoint 网页预览窗口）；如果不想保持 Excel 打开，"
        "请在设置里改用「由程序打开文档链接」模式。{}".format(
            "（{}）".format(detail[:200]) if detail else ""
        )
    )


def _normalize_workbook_name(name):
    return re.sub(r"\s+", "", str(name or "")).lower()


def _find_workbook(app, target_name):
    target = _normalize_workbook_name(target_name)
    workbooks = []
    for index in range(1, int(app.Workbooks.Count) + 1):
        workbook = app.Workbooks(index)
        name = str(workbook.Name or "")
        full_name = str(workbook.FullName or "")
        workbooks.append((name, full_name))
        normalized_name = _normalize_workbook_name(name)
        normalized_full = _normalize_workbook_name(full_name)
        if target and (normalized_name == target or target in normalized_full):
            return workbook
    if len(workbooks) == 1:
        return app.Workbooks(1)
    names = "、".join(name for name, _ in workbooks) or "无"
    raise RuntimeError(
        "Excel 中未找到「{}」。当前打开的工作簿: {}。"
        "请先在 Excel 中打开周报（SharePoint → 在桌面应用中打开），再重试。".format(
            target_name or "周报", names
        )
    )


def workbook_status(url, target_name=""):
    _require_com()
    target = target_name or workbook_name_from_url(url)
    try:
        app = _attach_excel(url)
    except RuntimeError as error:
        return {
            "excel_running": False,
            "target": target,
            "found": False,
            "workbook": None,
            "workbooks": [],
            "message": str(error),
        }
    try:
        workbooks = []
        found = None
        normalized_target = _normalize_workbook_name(target)
        for index in range(1, int(app.Workbooks.Count) + 1):
            workbook = app.Workbooks(index)
            name = str(workbook.Name or "")
            full_name = str(workbook.FullName or "")
            item = {
                "name": name,
                "full_name": full_name,
                "read_only": bool(workbook.ReadOnly),
            }
            workbooks.append(item)
            if normalized_target and (
                _normalize_workbook_name(name) == normalized_target
                or normalized_target in _normalize_workbook_name(full_name)
            ):
                found = item
        if found is None and len(workbooks) == 1:
            found = workbooks[0]
        return {
            "excel_running": True,
            "target": target,
            "found": found is not None,
            "workbook": found,
            "workbooks": workbooks,
            "message": "",
        }
    finally:
        pythoncom.CoUninitialize()


def _is_web_page_workbook(workbook):
    name = str(getattr(workbook, "Name", "") or "").lower()
    if name.endswith((".aspx", ".html", ".htm")):
        return True
    try:
        sheet_names = [
            str(sheet.Name).strip().lower() for sheet in workbook.Worksheets
        ]
    except Exception:
        return False
    web_names = {"doc", "doc2", "wopiframe", "wopi"}
    return bool(sheet_names) and all(name in web_names for name in sheet_names)


def _read_grid(worksheet, first_row, last_row, first_col, last_col):
    if last_row < first_row or last_col < first_col:
        return []
    values = worksheet.Range(
        worksheet.Cells(first_row, first_col), worksheet.Cells(last_row, last_col)
    ).Value
    if values is None:
        return []
    if not isinstance(values, tuple):
        return [[values]]
    if values and not isinstance(values[0], tuple):
        return [list(values)]
    return [list(row) for row in values]


def _open_workbook(app, url, read_only):
    try:
        workbook = app.Workbooks.Open(
            Filename=url,
            UpdateLinks=0,
            ReadOnly=read_only,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )
    except com_error:
        raise RuntimeError(
            "无法打开工作簿：文件可能已在 Excel 中打开或被占用。"
            "请先关闭 Excel 中已打开的该文件，或改用「附着到已打开的 Excel」模式。"
        )
    if not hasattr(workbook, "Worksheets"):
        raise RuntimeError(
            "工作簿未能正常打开（可能处于受保护视图），请在 Excel 中手动打开一次后再试"
        )
    if _is_web_page_workbook(workbook):
        try:
            workbook.Close(SaveChanges=False)
        except Exception:
            pass
        raise RuntimeError(
            "Excel 把该链接当成了网页而不是工作簿。请改用「附着到已打开的 Excel」模式："
            "先在 SharePoint 中点「在桌面应用中打开」，再回到本程序写入。"
        )
    return workbook


def _select_sheet(workbook, name):
    if name:
        try:
            return workbook.Worksheets(name)
        except com_error:
            available = [sheet.Name for sheet in workbook.Worksheets]
            raise RuntimeError(
                "找不到工作表「{}」，当前工作表: {}".format(name, "、".join(available))
            )
    return workbook.ActiveSheet


def _sheet_bounds(worksheet):
    used = worksheet.UsedRange
    first_row = int(used.Row)
    last_row = first_row + int(used.Rows.Count) - 1
    first_col = int(used.Column)
    last_col = first_col + int(used.Columns.Count) - 1
    return first_row, last_row, first_col, last_col


def _detect_header_com(worksheet, aliases, forced_row=0):
    first_row, last_row, first_col, last_col = _sheet_bounds(worksheet)
    scan_last_row = min(last_row, first_row + 39)
    scan_last_col = min(last_col, first_col + 79)
    grid = _read_grid(worksheet, first_row, scan_last_row, first_col, scan_last_col)
    alias_index = _alias_index(aliases)
    rows = []
    for offset, values in enumerate(grid):
        row_number = first_row + offset
        padded = [None] * (first_col - 1) + list(values)
        rows.append((row_number, padded))
    return _pick_best_header(rows, alias_index, forced_row)


def _find_data_bounds(worksheet, header_row, columns, footer_keywords):
    _, last_row, _, _ = _sheet_bounds(worksheet)
    ordered_columns = sorted(set(columns.values()))
    footer_columns = ordered_columns[:2]
    ticket_column = columns.get("ticket")
    if ticket_column and ticket_column not in footer_columns:
        footer_columns.append(ticket_column)
    data_columns = [
        columns[field] for field in ("ticket", "description") if field in columns
    ]
    if not data_columns:
        data_columns = ordered_columns[:2]
    last_data = header_row
    last_used = header_row
    footer = None
    for row in range(header_row + 1, last_row + 1):
        footer_texts = [
            str(worksheet.Cells(row, column).Value or "").strip()
            for column in footer_columns
        ]
        footer_text = None
        for text in footer_texts:
            if text and len(text) <= 16 and any(
                keyword and text.startswith(keyword) for keyword in footer_keywords
            ):
                footer_text = text
                break
        if footer_text:
            footer = row
            break
        data_texts = [
            str(worksheet.Cells(row, column).Value or "").strip()
            for column in data_columns
        ]
        if any(data_texts):
            last_data = row
        all_texts = [
            str(worksheet.Cells(row, column).Value or "").strip()
            for column in ordered_columns
        ]
        if any(all_texts):
            last_used = row
    return last_data, footer, last_used


def _to_number(value):
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip().replace("%", "").replace(",", "")
    if not text:
        return None
    try:
        return float(text) if "." in text else int(text)
    except ValueError:
        return None


def _same_value(written, expected):
    if isinstance(written, float) and written.is_integer():
        written = int(written)
    if isinstance(expected, float) and expected.is_integer():
        expected = int(expected)
    return str("" if written is None else written).strip() == str(
        "" if expected is None else expected
    ).strip()


def _set_text(cell, text):
    try:
        cell.NumberFormat = "@"
    except Exception:
        pass
    cell.Value = text


FORMULA_RANGE = re.compile(r"(\$?[A-Z]{1,3}\$?)(\d+):(\$?[A-Z]{1,3}\$?)(\d+)")


def _expand_footer_formulas(
    worksheet, footer_row, last_data_row, new_end_row, logger, warnings
):
    if not footer_row:
        return
    _, _, _, last_col = _sheet_bounds(worksheet)
    for column in range(1, last_col + 1):
        cell = worksheet.Cells(footer_row, column)
        try:
            if not cell.HasFormula:
                continue
            formula = cell.Formula
        except Exception:
            continue
        if not formula:
            continue

        def replace(match):
            start_col, start_row, end_col, end_row_text = match.groups()
            if int(end_row_text) == last_data_row:
                return "{}:{}{}".format(start_col + start_row, end_col, new_end_row)
            return match.group(0)

        updated = FORMULA_RANGE.sub(replace, formula)
        if updated != formula:
            try:
                cell.Formula = updated
                logger.info(
                    "已扩展页脚公式 {}: {} -> {}".format(
                        cell.Address, formula, updated
                    )
                )
            except Exception as error:
                warnings.append("页脚公式扩展失败: {}".format(error))
                logger.warn(warnings[-1])


def _row_value(row, field):
    if field == "repo":
        return row.get("repos")
    if field == "branch":
        return row.get("branches")
    if field == "date":
        return row.get("write_date", row.get("date_range"))
    if field == "commits":
        return row.get("commit_count")
    return row.get(field)


def write_rows(url, rows, options, logger=None, dry_run=True):
    logger = logger or NullLogger()
    _require_com()

    normalized_url = normalize_sharepoint_url(url)
    if normalized_url != str(url or "").strip():
        logger.info("已规范化文档链接: {}".format(normalized_url))
    url = normalized_url

    aliases = options.get("aliases") or {}
    footer_keywords = [
        str(keyword).strip()
        for keyword in (options.get("footer_keywords") or [])
        if str(keyword).strip()
    ]
    forced_header_row = int(options.get("header_row") or 0)
    sheet_name = (options.get("target_sheet") or "").strip()
    copy_row_format = bool(options.get("copy_row_format", True))
    insert_before_footer = bool(options.get("insert_before_footer", True))
    reuse_empty_rows = bool(options.get("reuse_empty_rows", True))
    snapshot_path = options.get("snapshot_path") or ""
    target_mode = (options.get("target_mode") or "existing").strip().lower()
    weekly_template_sheet = (options.get("weekly_template_sheet") or "").strip()
    weekly_sheet_prefix = options.get("weekly_sheet_prefix") or "周报汇总"
    weekly_overwrite = bool(options.get("weekly_overwrite", True))
    range_start = options.get("range_start") or ""

    write_mode = (options.get("write_mode") or "attach").strip().lower()
    target_workbook = options.get("target_workbook") or workbook_name_from_url(url)
    app = None
    workbook = None
    owns_app = False
    original_autosave = None
    original_display_alerts = None
    result = WriteResult(plan=WritePlan("", 0, 0, 0, {}))
    try:
        if write_mode == "attach":
            app = _attach_excel(url)
            logger.info("已附着到正在运行的 Excel")
            workbook = _find_workbook(app, target_workbook)
            result.read_only = bool(workbook.ReadOnly)
            logger.info(
                "已找到工作簿: {}（{}）".format(
                    workbook.Name, "只读" if result.read_only else "可写"
                )
            )
            if result.read_only and not dry_run:
                raise RuntimeError(
                    "工作簿是只读的，无法写入。请确认文件未被他人占用，"
                    "并在 Excel 中关闭只读提示后重试。"
                )
            try:
                original_display_alerts = app.DisplayAlerts
                app.DisplayAlerts = False
            except Exception:
                original_display_alerts = None
            can_write = not result.read_only and not dry_run
            if can_write:
                try:
                    original_autosave = bool(workbook.AutoSaveOn)
                    workbook.AutoSaveOn = False
                except Exception:
                    original_autosave = None
            if dry_run:
                logger.info("演练模式：只分析表头与插入位置，不改动已打开的工作簿")
        else:
            app = _start_excel()
            owns_app = True
            try:
                app.Calculation = XL_CALCULATION_MANUAL
            except Exception:
                pass
            logger.info("正在通过 Excel 打开在线工作簿…")
            workbook = _open_workbook(app, url, read_only=dry_run)
            result.read_only = bool(workbook.ReadOnly)
            logger.info(
                "已打开工作簿: {}（{}）".format(
                    workbook.Name, "只读" if result.read_only else "可写"
                )
            )
            if result.read_only and not dry_run:
                raise RuntimeError(
                    "工作簿以只读方式打开，无法写入。请确认：文件未被他人占用、"
                    "当前 Office 账号有编辑权限、且没有在浏览器里以只读模式打开。"
                )
            if result.read_only:
                logger.warn(
                    "工作簿以只读方式打开，演练只分析表头与插入位置，不会改动单元格"
                )
            can_write = not result.read_only
            if can_write:
                try:
                    workbook.AutoSaveOn = False
                except Exception as error:
                    if str(url).lower().startswith("http"):
                        result.warnings.append(
                            "无法关闭 Excel 自动保存: {}".format(error)
                        )
                        logger.warn(result.warnings[-1])

        if snapshot_path:
            try:
                Path(snapshot_path).parent.mkdir(parents=True, exist_ok=True)
                workbook.SaveCopyAs(str(Path(snapshot_path).resolve()))
                result.snapshot_path = str(Path(snapshot_path).resolve())
                logger.info("已保存写入前快照: {}".format(result.snapshot_path))
            except Exception as error:
                result.warnings.append("快照保存失败: {}".format(error))
                logger.warn(result.warnings[-1])

        plan_sheet_name = ""
        if target_mode == "weekly":
            template_name = weekly_template_sheet or _auto_weekly_template(workbook)
            if not template_name:
                raise RuntimeError(
                    "未找到周报汇总模板 sheet。请在设置里指定模板工作表。"
                )
            template_sheet = _select_sheet(workbook, template_name)
            new_name = weekly_sheet_name(weekly_sheet_prefix, range_start)
            if new_name == template_sheet.Name:
                raise RuntimeError(
                    "新 sheet 名称「{}」与模板相同，请修改前缀或模板。".format(
                        new_name
                    )
                )
            plan_sheet_name = new_name
            if not can_write:
                worksheet = template_sheet
                logger.info(
                    "演练模式：不创建新 sheet，仅分析模板「{}」，新 sheet 将为「{}」".format(
                        template_sheet.Name, new_name
                    )
                )
            else:
                existing_sheet = None
                try:
                    existing_sheet = workbook.Worksheets(new_name)
                except com_error:
                    existing_sheet = None
                if existing_sheet is not None and not weekly_overwrite:
                    base = new_name[:26]
                    for index in range(2, 100):
                        candidate = "{} ({})".format(base, index)
                        try:
                            workbook.Worksheets(candidate)
                        except com_error:
                            new_name = candidate
                            existing_sheet = None
                            break
                    if existing_sheet is not None:
                        raise RuntimeError("无法找到可用的新 sheet 名称")
                    plan_sheet_name = new_name
                if existing_sheet is not None:
                    worksheet = existing_sheet
                    logger.info(
                        "已存在 sheet「{}」，将清空数据后重写".format(new_name)
                    )
                else:
                    template_sheet.Copy(
                        None, workbook.Worksheets(workbook.Worksheets.Count)
                    )
                    worksheet = workbook.Worksheets(workbook.Worksheets.Count)
                    worksheet.Name = new_name
                    logger.info(
                        "已从模板「{}」创建新 sheet「{}」".format(
                            template_sheet.Name, new_name
                        )
                    )
        else:
            worksheet = _select_sheet(workbook, sheet_name)
            plan_sheet_name = worksheet.Name

        header_row, columns = _detect_header_com(worksheet, aliases, forced_header_row)
        if not header_row or not columns:
            raise RuntimeError(
                "未能在工作表「{}」中识别表头。请在设置里指定表头行，或检查字段别名。".format(
                    worksheet.Name
                )
            )
        missing = [field for field in REQUIRED_FIELDS if field not in columns]
        if missing:
            raise RuntimeError(
                "表头缺少必要字段: {}。识别到的列: {}".format(
                    "、".join(missing),
                    "、".join(
                        "{}={}".format(field, column)
                        for field, column in sorted(columns.items(), key=lambda kv: kv[1])
                    ),
                )
            )
        for field in FIELD_ORDER:
            if field not in columns:
                logger.info("表头未找到字段「{}」，该字段将跳过".format(field))

        last_data, footer, last_used = _find_data_bounds(
            worksheet, header_row, columns, footer_keywords
        )
        logger.info(
            "表头第 {} 行，最后数据行 {}，最后使用行 {}，页脚行 {}".format(
                header_row, last_data, last_used, footer if footer else "无"
            )
        )

        format_source_row = None
        if target_mode == "weekly":
            format_source_row = (
                header_row + 1 if last_data >= header_row + 1 else None
            )
            if can_write and last_data >= header_row + 1:
                worksheet.Rows(
                    "{0}:{1}".format(header_row + 1, last_data)
                ).ClearContents()
                logger.info(
                    "已清空模板数据行 {}-{}".format(header_row + 1, last_data)
                )
            start_row = header_row + 1
            end_row = start_row + len(rows) - 1
            final_footer_row = None
            plan = WritePlan(
                sheet_name=plan_sheet_name,
                header_row=header_row,
                start_row=start_row,
                end_row=end_row,
                columns=columns,
                footer_row=None,
                last_data_row=header_row,
            )
        else:
            if footer and insert_before_footer:
                start_row = footer
            elif footer:
                start_row = footer + 1
            elif reuse_empty_rows:
                start_row = last_data + 1
                if last_used > last_data:
                    logger.info(
                        "将从第 {} 行开始，复用未填充的日期行".format(start_row)
                    )
            else:
                start_row = last_used + 1
                if last_used > last_data:
                    logger.info(
                        "将追加到第 {} 行，保留未填充的日期行".format(start_row)
                    )
            end_row = start_row + len(rows) - 1
            final_footer_row = (
                footer + len(rows) if footer and insert_before_footer else footer
            )
            plan = WritePlan(
                sheet_name=plan_sheet_name,
                header_row=header_row,
                start_row=start_row,
                end_row=end_row,
                columns=columns,
                footer_row=final_footer_row,
                last_data_row=last_data,
            )
        result.plan = plan
        logger.info(
            "准备写入 {} 行：工作表「{}」第 {}~{} 行".format(
                len(rows), plan_sheet_name, start_row, end_row
            )
        )

        if can_write and target_mode != "weekly" and footer and len(rows) > 0:
            worksheet.Rows(
                "{0}:{1}".format(start_row, start_row + len(rows) - 1)
            ).Insert(Shift=XL_SHIFT_DOWN)
            logger.info(
                "已在页脚{}插入 {} 行".format(
                    "前" if insert_before_footer else "后", len(rows)
                )
            )

        format_row = format_source_row if target_mode == "weekly" else last_data
        if (
            can_write
            and copy_row_format
            and format_row
            and format_row >= header_row + 1
            and len(rows) > 0
        ):
            try:
                worksheet.Rows(format_row).Copy()
                worksheet.Rows(
                    "{0}:{1}".format(start_row, end_row)
                ).PasteSpecial(Paste=XL_PASTE_FORMATS)
                app.CutCopyMode = False
                logger.info("已复用第 {} 行的格式".format(format_row))
            except Exception as error:
                app.CutCopyMode = False
                result.warnings.append("复用行格式失败: {}".format(error))
                logger.warn(result.warnings[-1])

        if can_write:
            for offset, row in enumerate(rows):
                excel_row = start_row + offset
                for field, column in columns.items():
                    if field not in FIELD_ORDER:
                        continue
                    value = _row_value(row, field)
                    if field in ("hours", "progress"):
                        number = _to_number(value)
                        if number is None:
                            continue
                        worksheet.Cells(excel_row, column).Value = number
                        continue
                    if field == "commits":
                        if value in (None, ""):
                            continue
                        worksheet.Cells(excel_row, column).Value = int(value)
                        continue
                    if field == "date":
                        if isinstance(value, datetime):
                            value = value.date()
                        if isinstance(value, date):
                            serial = (value - date(1899, 12, 30)).days
                            worksheet.Cells(excel_row, column).Value2 = float(serial)
                            continue
                    text = "" if value is None else str(value)
                    if text.startswith("="):
                        text = "'" + text
                        result.escaped += 1
                    cell = worksheet.Cells(excel_row, column)
                    _set_text(cell, text)
                    if field == "description":
                        cell.WrapText = True

        if can_write and target_mode != "weekly" and footer and insert_before_footer:
            _expand_footer_formulas(
                worksheet,
                final_footer_row,
                last_data,
                end_row,
                logger,
                result.warnings,
            )

        first_ticket = _row_value(rows[0], "ticket") if rows else ""
        last_ticket = _row_value(rows[-1], "ticket") if rows else ""
        ticket_column = columns.get("ticket")
        if can_write and ticket_column and rows:
            written_first = worksheet.Cells(start_row, ticket_column).Value
            written_last = worksheet.Cells(end_row, ticket_column).Value
            result.verified = _same_value(
                written_first, first_ticket
            ) and _same_value(written_last, last_ticket)
            if not result.verified:
                result.warnings.append("写入后回读校验不一致，请打开文件确认")
                logger.warn(result.warnings[-1])

        if not can_write:
            logger.info("演练模式：未保存工作簿")
        else:
            workbook.Save()
            result.saved = True
            if owns_app:
                logger.info("已保存并同步到 SharePoint")
            else:
                logger.info("已保存，Excel 中的工作簿已更新")

        return result
    except Exception:
        logger.error("写入失败:\n{}".format(traceback.format_exc()))
        raise
    finally:
        if owns_app:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            _stop_excel(app)
        else:
            if workbook is not None:
                if original_autosave is not None:
                    try:
                        workbook.AutoSaveOn = original_autosave
                    except Exception:
                        pass
                if original_display_alerts is not None:
                    try:
                        app.DisplayAlerts = original_display_alerts
                    except Exception:
                        pass
            if app is not None:
                pythoncom.CoUninitialize()


def inspect_sharepoint(
    url,
    aliases,
    sheet_name="",
    forced_header_row=0,
    preview_rows=10,
    write_mode="attach",
    target_workbook="",
):
    _require_com()
    owns_app = False
    app = None
    workbook = None
    try:
        if (write_mode or "attach").strip().lower() == "attach":
            app = _attach_excel(url)
            workbook = _find_workbook(
                app, target_workbook or workbook_name_from_url(url)
            )
        else:
            app = _start_excel()
            owns_app = True
            workbook = _open_workbook(
                app, normalize_sharepoint_url(url), read_only=True
            )
        sheets = []
        detected = None
        for worksheet in workbook.Worksheets:
            first_row, last_row, first_col, last_col = _sheet_bounds(worksheet)
            scan_last_row = min(last_row, first_row + 39)
            scan_last_col = min(last_col, first_col + 79)
            grid = _read_grid(
                worksheet, first_row, scan_last_row, first_col, scan_last_col
            )
            alias_index = _alias_index(aliases)
            rows = []
            for offset, values in enumerate(grid):
                padded = [None] * (first_col - 1) + list(values)
                rows.append((first_row + offset, padded))
            header_row, columns = _pick_best_header(
                rows, alias_index, forced_header_row
            )
            preview = []
            preview_grid = _read_grid(
                worksheet,
                first_row,
                min(last_row, first_row + preview_rows - 1),
                first_col,
                min(last_col, first_col + 29),
            )
            for values in preview_grid:
                preview.append([_text_preview(value, 80) for value in values])
            item = {
                "name": worksheet.Name,
                "max_row": last_row,
                "max_column": last_col,
                "header_row": header_row,
                "columns": columns,
                "preview": preview,
            }
            sheets.append(item)
            if header_row and (
                detected is None or len(columns) > len(detected["columns"])
            ):
                detected = item
        return {
            "path": url,
            "read_only": bool(workbook.ReadOnly),
            "sheets": sheets,
            "detected": detected,
        }
    finally:
        if owns_app:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            _stop_excel(app)
        elif app is not None:
            pythoncom.CoUninitialize()
