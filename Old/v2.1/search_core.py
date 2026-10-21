# -*- coding: utf-8 -*-
"""
search_core.py — Excel / CSV 關鍵字搜尋核心

不依賴 PyQt,可被呼叫,也可直接用命令列執行:
    python search_core.py "C:\\Users\\User\\Desktop" result.xlsx -k "google chrome" -k "microsoft edge"

需要套件:  pip install pandas numpy openpyxl xlrd
"""
from __future__ import annotations

import multiprocessing
import os
import re
import time
import unicodedata
from collections import Counter
from concurrent.futures import FIRST_COMPLETED, ProcessPoolExecutor, wait
from concurrent.futures.process import BrokenProcessPool
from dataclasses import dataclass, field
from typing import Callable, Iterator, Optional

import numpy as np
import pandas as pd
from openpyxl.utils import get_column_letter

ALL_EXTS = (".csv", ".xlsx", ".xls")
EXCEL_MAX_ROWS = 1_048_575          # 扣掉標題列
CSV_CHUNK_ROWS = 200_000            # 大型 CSV 分塊讀取,避免一次吃光記憶體
POOL_MIN_FILES = 6                  # 檔案數少於此值就不開多核心 (啟動成本不划算)
HEADER_SCAN_ROWS = 30               # 只在工作表最上方這麼多列裡找標題列
META_COLUMNS = ("檔案路徑", "檔案名稱", "來源工作表", "來源列號", "命中關鍵字")
_RESERVED = META_COLUMNS + ("資料列序",)   # 原始資料若有同名欄位,會加上「_原始」
_SEP = "\x1f"                       # 串接整列時的分隔字元,避免關鍵字跨儲存格誤判
_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
_MIDNIGHT = re.compile(r"^(\d{4}-\d{2}-\d{2}) 00:00:00$")
# 長得像「資料」而不是「欄位名稱」的儲存格:數字、版本號 (154.0.8037.98)、日期
_DATA_LIKE = re.compile(
    r"^(?:[+-]?\d+(?:[.,]\d+)*"
    r"|\d{4}[-/.]\d{1,2}[-/.]\d{1,2}(?:[ T].*)?"
    r"|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}(?:[ T].*)?)$"
)


# ───────────────────────────── 資料結構 ─────────────────────────────
class SearchError(Exception):
    """可直接顯示給使用者看的錯誤。"""


@dataclass(frozen=True)
class SearchOptions:
    keywords: tuple
    case_sensitive: bool = False
    whole: bool = False
    whole_mode: str = "cell"        # "cell": 整個儲存格相同 / "word": 完整單字
    match_all: bool = False         # True: 全部關鍵字都要出現;False: 任一個即可
    normalize_width: bool = False   # True: 全形半形視為相同
    exts: tuple = ALL_EXTS
    recursive: bool = True
    skip_hidden: bool = True
    auto_header: bool = True        # 自動判斷 Excel 工作表的標題列 (略過上方的標題文字、空白列)
    merge_columns: bool = True      # 合併名稱相近、或沒有標題的欄位
    aliases: tuple = ()             # 欄位別名,每項寫成 "版本=軟體版本"


@dataclass
class Part:
    """一張工作表 (或 CSV 的一個區塊) 中符合的資料。"""
    frame: pd.DataFrame             # 前 5 欄是 META_COLUMNS,後面是原始資料欄位
    cols: list                      # 原始資料欄位名稱 (依序)
    sheet: str
    generic: bool                   # True = 找不到標題列,欄位以 欄A、欄B… 命名
    header_row: Optional[int]       # 標題列在 Excel 中的列號 (1 開始);None = 沒有標題


@dataclass
class FileResult:
    path: str
    parts: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    error: Optional[str] = None


@dataclass
class SearchSummary:
    cancelled: bool
    scanned: int
    matched_files: int
    rows: int
    elapsed: float
    output: Optional[str] = None
    preview: Optional[pd.DataFrame] = None


@dataclass
class Sheet:
    name: Optional[str]
    df: pd.DataFrame
    rows: np.ndarray                # 每一列在原檔案中的列號 (Excel 開啟時看到的列號)
    generic: bool
    header_row: Optional[int]
    header_cells: list = field(default_factory=list)   # 標題列的原始內容 (找不到標題時為空)


# ───────────────────────────── 檔案蒐集 ─────────────────────────────
def _is_hidden(full: str, name: str) -> bool:
    if name.startswith("."):
        return True
    if os.name == "nt":
        try:
            return bool(os.stat(full).st_file_attributes & 2)  # FILE_ATTRIBUTE_HIDDEN
        except (OSError, AttributeError):
            return False
    return False


def collect_files(source: str, opts: SearchOptions, exclude: Optional[str] = None) -> list:
    """列出要處理的檔案。exclude 通常是輸出檔,避免把自己的結果又掃進去。"""
    exts = {e.lower() for e in opts.exts}
    ex = os.path.normcase(os.path.abspath(exclude)) if exclude else None
    files = []
    for root, dirs, names in os.walk(source):
        if opts.skip_hidden:
            dirs[:] = [d for d in dirs if not _is_hidden(os.path.join(root, d), d)]
        dirs.sort()
        for n in sorted(names):
            if n.startswith("~$"):  # Excel 暫存檔
                continue
            if os.path.splitext(n)[1].lower() not in exts:
                continue
            full = os.path.join(root, n)
            if opts.skip_hidden and _is_hidden(full, n):
                continue
            if ex and os.path.normcase(os.path.abspath(full)) == ex:
                continue
            files.append(full)
        if not opts.recursive:
            break
    return files


# ───────────────────────────── 標題列與欄位名稱 ─────────────────────────────
_CJK_GAP = re.compile(r"(?<=[\u3400-\u9fff])\s+(?=[\u3400-\u9fff])")


def normalize_header(s) -> str:
    """欄位名稱正規化:全形轉半形、連續空白(含換行)合併、中文字之間的空白去掉、去頭尾空白。"""
    s = re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(s))).strip()
    return _CJK_GAP.sub("", s)


def parse_aliases(lines) -> dict:
    """把 ["版本=軟體版本", ...] 轉成 {正規化後的別名(小寫): 標準名稱}。"""
    out = {}
    for line in lines or ():
        parts = re.split(r"\s*(?:=|＝|->|→)\s*", str(line).strip(), maxsplit=1)
        if len(parts) == 2 and normalize_header(parts[0]) and normalize_header(parts[1]):
            out[normalize_header(parts[0]).casefold()] = normalize_header(parts[1])
    return out


def build_columns(names: list, aliases: dict) -> list:
    """產生最終欄位名稱:正規化 → 套用別名 → 空白標題改叫「欄C」→ 重複的加 .2"""
    out, seen = [], {}
    for i, n in enumerate(names):
        n = normalize_header(n)
        n = aliases.get(n.casefold(), n)
        if not n:
            n = f"欄{get_column_letter(i + 1)}"
        if n in seen:
            seen[n] += 1
            n = f"{n}.{seen[n]}"
        else:
            seen[n] = 0
        out.append(n)
    return out


def find_header_row(raw: pd.DataFrame, max_scan: int = HEADER_SCAN_ROWS) -> Optional[int]:
    """在最上方幾列找出「真正的標題列」,回傳 0 起算的位置;找不到回傳 None。

    標題列的特徵:
      1. 有填值的儲存格數量接近整張表的一般寬度 (合併儲存格的大標題通常只有 1 格,會被略過)
      2. 內容像文字,不像數字、版本號或日期 (所以資料列不會被誤當成標題)
    """
    head = raw.iloc[:max_scan]
    if head.empty:
        return None
    filled = head.apply(lambda c: c.str.strip().ne("")).to_numpy()
    counts = filled.sum(axis=1)
    nonzero = counts[counts > 0]
    if nonzero.size == 0:
        return None
    ref = int(np.ceil(np.percentile(nonzero, 75)))
    need = 1 if ref <= 1 else max(2, int(np.ceil(ref * 0.6)))
    for i in range(len(head)):
        if counts[i] < need:
            continue
        cells = [v.strip() for v in head.iloc[i] if v.strip()]
        if any(_DATA_LIKE.match(v) for v in cells):
            continue
        return i
    return None


def _tidy_dates(raw: pd.DataFrame) -> pd.DataFrame:
    """Excel 的日期儲存格讀進來會變成「2026-10-02 00:00:00」,把沒有時間的去掉 00:00:00。"""
    for c in raw.columns:
        s = raw[c]
        if s.str.contains(" 00:00:00", regex=False).any():
            raw[c] = s.str.replace(_MIDNIGHT, r"\1", regex=True)
    return raw


def prepare_sheet(name: Optional[str], raw: pd.DataFrame, opts: SearchOptions, aliases: dict,
                  force_generic: bool = False) -> Sheet:
    """把「沒有標題、從第 1 列開始」的原始表格,整理成有欄位名稱的資料表。"""
    raw = _tidy_dates(raw.fillna("").astype(str))
    h = None if force_generic else (find_header_row(raw) if opts.auto_header else 0)
    generic = h is None
    names = [""] * raw.shape[1] if generic else raw.iloc[h].tolist()
    body = raw if generic else raw.iloc[h + 1:]
    cols = build_columns(names, aliases)
    # 沒有標題、而且整欄都是空白的欄位 (排版殘留) 直接丟掉
    keep = [j for j in range(len(cols))
            if normalize_header(names[j]) or body.iloc[:, j].str.strip().ne("").any()]
    body = body.iloc[:, keep].copy()
    body.columns = [cols[j] for j in keep]
    rows = body.index.to_numpy() + 1        # raw 的列位置 0 起算 → Excel 列號
    return Sheet(name, body, rows, generic, None if generic else h + 1,
                 [] if generic else [str(v) for v in raw.iloc[h].tolist()])


def _header_has_hit(cells: list, opts: SearchOptions) -> bool:
    idx, _ = match_frame(pd.DataFrame([cells]), opts)
    return len(idx) > 0


# ───────────────────────────── 讀檔 ─────────────────────────────
def detect_csv_encoding(path: str) -> str:
    """只解碼不解析,先找出能完整讀完整個檔案的編碼。"""
    with open(path, "rb") as f:
        head = f.read(4)
    if head[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return "utf-16"
    for enc in ("utf-8-sig", "cp950", "gb18030"):
        try:
            with open(path, "r", encoding=enc, newline="") as f:
                while f.read(1 << 20):
                    pass
            return enc
        except UnicodeDecodeError:
            continue
    return "latin-1"


def read_sheets(path: str, opts: SearchOptions, aliases: dict) -> Iterator[Sheet]:
    """逐塊產生 Sheet。所有欄位都以字串讀入。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        enc = detect_csv_encoding(path)
        try:
            with pd.read_csv(path, dtype=str, encoding=enc, na_filter=False,
                             chunksize=CSV_CHUNK_ROWS) as reader:
                for chunk in reader:
                    names = ["" if str(c).startswith("Unnamed: ") else c for c in chunk.columns]
                    chunk.columns = build_columns(names, aliases)
                    # 檔案第 1 行是標題,所以資料從第 2 行開始 (和用 Excel 開啟時看到的列號一致)
                    yield Sheet(None, chunk, chunk.index.to_numpy() + 2, False, 1)
        except pd.errors.EmptyDataError:
            return
    else:
        with pd.ExcelFile(path) as xl:
            for sheet in xl.sheet_names:
                raw = pd.read_excel(xl, sheet_name=sheet, header=None, dtype=str, na_filter=False)
                if raw.empty:
                    continue
                sh = prepare_sheet(sheet, raw, opts, aliases)
                if sh.header_cells and _header_has_hit(sh.header_cells, opts):
                    # 被判成標題的那一列其實含有關鍵字,多半是沒有標題的資料 → 保留成資料,避免漏掉
                    sh = prepare_sheet(sheet, raw, opts, aliases, force_generic=True)
                yield sh


# ───────────────────────────── 比對 ─────────────────────────────
def _transform(s: pd.Series, opts: SearchOptions) -> pd.Series:
    if opts.normalize_width:
        s = s.str.normalize("NFKC")
    if not opts.case_sensitive:
        s = s.str.casefold()
    return s


def _transform_kw(kw: str, opts: SearchOptions) -> str:
    if opts.normalize_width:
        kw = unicodedata.normalize("NFKC", kw)
    if not opts.case_sensitive:
        kw = kw.casefold()
    return kw


def clean_keywords(opts: SearchOptions) -> list:
    kws = [str(k) for k in opts.keywords if str(k).strip()]
    if not kws:
        raise SearchError("請至少提供一個搜尋關鍵字。")
    return kws


def match_frame(df: pd.DataFrame, opts: SearchOptions):
    """回傳 (符合列的位置陣列, 每列命中的關鍵字文字)。"""
    keywords = clean_keywords(opts)
    clean = df.fillna("").astype(str)
    if clean.shape[0] == 0 or clean.shape[1] == 0:
        return np.array([], dtype=int), []

    masks = []
    if opts.whole and opts.whole_mode == "cell":
        # 任一儲存格的內容 (去頭尾空白後) 必須恰好等於關鍵字
        cells = clean.apply(lambda c: _transform(c.str.strip(), opts))
        for kw in keywords:
            k = _transform_kw(kw.strip(), opts)
            masks.append(cells.eq(k).to_numpy().any(axis=1))
    else:
        # 把整列串成一段文字再搜尋;分隔字元不是空白,所以關鍵字不會跨儲存格誤判
        text = clean.iloc[:, 0]
        for i in range(1, clean.shape[1]):
            text = text + _SEP + clean.iloc[:, i]
        text = _transform(text, opts)
        for kw in keywords:
            k = _transform_kw(kw, opts)
            if opts.whole:  # 完整單字:前後不可緊接文字
                pat = rf"(?<!\w){re.escape(k)}(?!\w)"
                masks.append(text.str.contains(pat, regex=True).to_numpy())
            else:
                masks.append(text.str.contains(k, regex=False).to_numpy())

    arr = np.column_stack(masks)
    hit = arr.all(axis=1) if opts.match_all else arr.any(axis=1)
    idx = np.flatnonzero(hit)
    labels = [", ".join(keywords[j] for j in np.flatnonzero(arr[i])) for i in idx]
    return idx, labels


def process_file(path: str, opts: SearchOptions) -> FileResult:
    """處理單一檔案 (會在子行程中執行,所以必須是模組層級函式)。"""
    try:
        name = os.path.basename(path)
        aliases = parse_aliases(opts.aliases)
        parts = []
        for sh in read_sheets(path, opts, aliases):
            df = sh.df
            if df.empty or df.shape[1] == 0:
                continue
            idx, labels = match_frame(df, opts)
            if len(idx) == 0:
                continue
            part = df.iloc[idx].copy()
            part.columns = [f"{c}_原始" if c in _RESERVED else c for c in part.columns]
            part.reset_index(drop=True, inplace=True)
            meta = pd.DataFrame({
                "檔案路徑": path,
                "檔案名稱": name,
                "來源工作表": sh.name or "",
                "來源列號": sh.rows[idx],
                "命中關鍵字": labels,
            })
            parts.append(Part(pd.concat([meta, part], axis=1), list(part.columns),
                              sh.name or "", sh.generic, sh.header_row))
        return FileResult(path, parts, _header_notes(name, parts))
    except Exception as e:  # noqa: BLE001
        return FileResult(path, error=str(e) or type(e).__name__)


def _header_notes(name: str, parts: list) -> list:
    """標題列不在第 1 列、或根本找不到標題時,留下一行說明,方便檢查。"""
    moved = [p for p in parts if p.header_row not in (None, 1) and p.sheet]
    nohdr = [p for p in parts if p.generic]
    notes = []
    if moved:
        rows = sorted({p.header_row for p in moved})
        notes.append(f"{name}:{len(moved)} 張工作表的標題列不在第 1 列 (在第 {'、'.join(map(str, rows))} 列),已自動判斷")
    if nohdr:
        notes.append(f"{name}:{len(nohdr)} 張工作表找不到標題列 ({'、'.join(p.sheet for p in nohdr[:4])}"
                     f"{'…' if len(nohdr) > 4 else ''})")
    return notes


# ───────────────────────────── 欄位整理 ─────────────────────────────
def reconcile_columns(parts: list) -> list:
    """把「格式不太一樣」的工作表對齊到主流格式。就地修改 parts,回傳要顯示的訊息。

    只處理兩種有把握的情況:
      1. 找不到標題列的工作表 → 如果有欄位數相同的主流格式,依欄位位置套用它的標題
      2. 某一欄的標題剛好等於工作表名稱 (常見於標題被誤改成工作表名稱),
         而其餘欄位和主流格式完全相同 → 視為主流格式中同一位置的那一欄
    其他情況不亂猜,請用「欄位別名」自行指定。
    """
    weight: Counter = Counter()     # 格式 → 命中筆數
    count: Counter = Counter()      # 格式 → 出現在幾張工作表
    for p in parts:
        if not p.generic:
            weight[tuple(p.cols)] += len(p.frame)
            count[tuple(p.cols)] += 1
    ranked = [s for s, _ in weight.most_common()]
    messages, seen = [], set()

    def dominant(width: int):
        """同樣欄數的格式中,要明顯佔多數 (至少 3 張工作表、且過半) 才算主流格式。"""
        same = [s for s in ranked if len(s) == width]
        if not same:
            return None
        top = max(same, key=lambda s: count[s])
        total = sum(count[s] for s in same)
        return top if count[top] >= 3 and count[top] * 2 >= total else None

    for p in parts:
        cols = tuple(p.cols)
        mapping, why = {}, ""
        if p.generic:
            top = dominant(len(cols))
            if top:
                mapping = dict(zip(cols, top))
                why = f"工作表「{p.sheet}」沒有標題列,依欄位位置套用:{' / '.join(top)}"
        elif len(cols) >= 3:
            for s in ranked:
                if s == cols or len(s) != len(cols) or weight[s] <= weight[cols]:
                    continue
                diff = [i for i, (a, b) in enumerate(zip(cols, s)) if a != b]
                if (len(diff) == 1 and s[diff[0]] not in cols
                        and normalize_header(cols[diff[0]]).casefold() == normalize_header(p.sheet).casefold()):
                    mapping = {cols[diff[0]]: s[diff[0]]}
                    why = f"工作表「{p.sheet}」的欄位標題「{cols[diff[0]]}」與工作表名稱相同,視為「{s[diff[0]]}」"
                    break
        if mapping:
            p.frame.columns = [mapping.get(c, c) if c not in META_COLUMNS else c for c in p.frame.columns]
            p.cols = [mapping.get(c, c) for c in p.cols]
            if why and why not in seen:
                seen.add(why)
                messages.append(why)
    return messages


def ordered_columns(parts: list) -> list:
    """資料欄位的排列順序:主流格式的欄位排前面。"""
    weight: Counter = Counter()
    for p in parts:
        weight[tuple(p.cols)] += len(p.frame)
    out: list = []
    for cols, _ in weight.most_common():
        out.extend(c for c in cols if c not in out)
    return out


# ───────────────────────────── 輸出 ─────────────────────────────
def with_timestamp(path: str) -> str:
    root, ext = os.path.splitext(path)
    return f"{root}_{time.strftime('%Y%m%d_%H%M%S')}{ext}"


def _unique(names: list) -> list:
    seen, out = {}, []
    for n in map(str, names):
        if n in seen:
            seen[n] += 1
            out.append(f"{n}.{seen[n]}")
        else:
            seen[n] = 0
            out.append(n)
    return out


def _display_width(s: str) -> int:
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in s)


def write_output(final: pd.DataFrame, out: str, log: Callable = lambda *_: None) -> str:
    """寫出結果,回傳實際寫入的路徑 (超過 Excel 列數上限時會改存 CSV)。"""
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    final = final.copy()
    final.columns = _unique(list(final.columns))

    if out.lower().endswith(".xlsx") and len(final) > EXCEL_MAX_ROWS:
        out = os.path.splitext(out)[0] + ".csv"
        log("warn", f"結果超過 Excel 列數上限 ({EXCEL_MAX_ROWS:,} 列),改存為 CSV:{out}")

    if out.lower().endswith(".csv"):
        final.to_csv(out, index=False, encoding="utf-8-sig")
        return out

    for c in final.columns:  # openpyxl 不接受控制字元
        try:
            final[c] = final[c].str.replace(_ILLEGAL_XML, "", regex=True)
        except AttributeError:
            pass

    from openpyxl.styles import Alignment, Font, PatternFill

    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        final.to_excel(xw, index=False, sheet_name="結果")
        ws = xw.sheets["結果"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        head_font = Font(bold=True, color="FFFFFF")
        head_fill = PatternFill("solid", fgColor="4A66E8")
        for cell in ws[1]:
            cell.font = head_font
            cell.fill = head_fill
            cell.alignment = Alignment(vertical="center")
        ws.row_dimensions[1].height = 22
        for i, col in enumerate(final.columns, 1):
            sample = final[col].head(200).fillna("").astype(str)
            w = max([_display_width(str(col))] + [_display_width(v) for v in sample])
            ws.column_dimensions[get_column_letter(i)].width = min(max(w + 2, 8), 60)
    return out


# ───────────────────────────── 主流程 ─────────────────────────────
def _noop(*_a, **_k):
    return None


def run_search(
    source: str,
    output: str,
    opts: SearchOptions,
    *,
    workers: int = 1,
    on_progress: Callable = _noop,   # (已完成數, 總數, 目前檔案)
    on_stats: Callable = _noop,      # (已掃描, 命中檔案, 命中筆數)
    on_log: Callable = _noop,        # (等級, 訊息)
    should_cancel: Callable = lambda: False,
) -> SearchSummary:
    t0 = time.time()
    clean_keywords(opts)  # 提早驗證
    files = collect_files(source, opts, exclude=output)
    total = len(files)
    on_progress(0, total, "")

    def summary(**kw):
        base = dict(cancelled=False, scanned=0, matched_files=0, rows=0, elapsed=time.time() - t0)
        base.update(kw)
        return SearchSummary(**base)

    if total == 0:
        on_log("warn", "資料夾內沒有符合副檔名的檔案。")
        return summary()

    use_pool = workers > 1 and total >= POOL_MIN_FILES
    on_log("info", f"找到 {total} 個檔案,開始搜尋" + (f" (使用 {min(workers, total)} 個行程)…" if use_pool else "…"))

    results: dict = {}
    state = dict(scanned=0, files=0, rows=0)

    def handle(i: int, res: FileResult):
        results[i] = res
        state["scanned"] += 1
        if res.error:
            on_log("err", f"讀取失敗  {res.path}: {res.error}")
        elif res.parts:
            n = sum(len(p.frame) for p in res.parts)
            state["files"] += 1
            state["rows"] += n
            on_log("ok", f"找到符合項目  {res.path}  ({n} 筆)")
            for note in res.notes:
                on_log("info", note)
        on_stats(state["scanned"], state["files"], state["rows"])
        on_progress(state["scanned"], total, res.path)

    def run_sequential(indices):
        for i in indices:
            if should_cancel():
                return
            on_progress(state["scanned"], total, files[i])
            handle(i, process_file(files[i], opts))

    if use_pool:
        try:
            _run_pool(files, opts, workers, should_cancel, handle)
        except BrokenProcessPool:
            on_log("warn", "多核心處理失敗,剩餘檔案改用單核心繼續。")
            run_sequential([i for i in range(total) if i not in results])
    else:
        run_sequential(range(total))

    on_progress(state["scanned"], total, "")
    if should_cancel():
        on_log("warn", "已取消,未寫入結果。")
        return summary(cancelled=True, scanned=state["scanned"],
                       matched_files=state["files"], rows=state["rows"])

    parts = [p for i in sorted(results) for p in results[i].parts]
    if not parts:
        on_log("warn", "未找到包含指定關鍵字的資料。")
        return summary(scanned=state["scanned"])

    if opts.merge_columns:
        for msg in reconcile_columns(parts):
            on_log("info", msg)
    final = pd.concat([p.frame for p in parts], ignore_index=True, sort=False)
    final = final[list(META_COLUMNS) + ordered_columns(parts)]
    try:
        out = write_output(final, output, on_log)
    except PermissionError:
        raise SearchError(f"無法寫入「{output}」。\n檔案可能正被 Excel 開啟,請先關閉後再試。") from None
    except OSError as e:
        raise SearchError(f"寫入結果失敗:{e}") from None
    on_log("ok", f"搜尋完成,結果已寫入:{out}")
    return summary(scanned=state["scanned"], matched_files=state["files"], rows=state["rows"],
                   output=out, preview=final.head(500).fillna(""))


def _run_pool(files, opts, workers, should_cancel, handle):
    # 一律用 spawn:跟 Windows 行為一致,也避免在有 Qt 執行緒的行程裡 fork
    ctx = multiprocessing.get_context("spawn")
    ex = ProcessPoolExecutor(max_workers=min(workers, len(files)), mp_context=ctx)
    futs = {ex.submit(process_file, p, opts): i for i, p in enumerate(files)}
    pending = set(futs)
    try:
        while pending:
            if should_cancel():
                break
            done, pending = wait(pending, timeout=0.2, return_when=FIRST_COMPLETED)
            for f in done:
                handle(futs[f], f.result())
    finally:
        ex.shutdown(wait=True, cancel_futures=True)


# ───────────────────────────── 命令列 ─────────────────────────────
def _cli():
    import argparse

    ap = argparse.ArgumentParser(description="搜尋資料夾內 Excel / CSV 中含關鍵字的整列資料")
    ap.add_argument("source", help="來源資料夾")
    ap.add_argument("output", help="輸出檔案 (.xlsx 或 .csv)")
    ap.add_argument("-k", "--keyword", action="append", required=True, help="關鍵字,可重複指定")
    ap.add_argument("--case", action="store_true", help="區分大小寫")
    ap.add_argument("--whole", choices=["cell", "word"], help="完整比對方式")
    ap.add_argument("--all", action="store_true", help="全部關鍵字都要出現")
    ap.add_argument("--width", action="store_true", help="全形半形視為相同")
    ap.add_argument("--no-recursive", action="store_true", help="不搜尋子資料夾")
    ap.add_argument("--no-auto-header", action="store_true", help="一律把第 1 列當標題列")
    ap.add_argument("--no-merge", action="store_true", help="不合併相近或沒有標題的欄位")
    ap.add_argument("-a", "--alias", action="append", default=[], help='欄位別名,例如 -a "版本=軟體版本"')
    ap.add_argument("-j", "--jobs", type=int, default=0, help="行程數 (預設自動)")
    a = ap.parse_args()

    opts = SearchOptions(
        keywords=tuple(a.keyword), case_sensitive=a.case, whole=bool(a.whole),
        whole_mode=a.whole or "cell", match_all=a.all, normalize_width=a.width,
        recursive=not a.no_recursive, auto_header=not a.no_auto_header,
        merge_columns=not a.no_merge, aliases=tuple(a.alias))
    jobs = a.jobs or max(1, min(8, (os.cpu_count() or 2) - 1))
    try:
        r = run_search(a.source, a.output, opts, workers=jobs,
                       on_log=lambda lv, m: print(f"[{lv}] {m}"))
    except SearchError as e:
        raise SystemExit(str(e))
    print(f"掃描 {r.scanned} 個檔案,命中 {r.matched_files} 個檔案、{r.rows} 筆,耗時 {r.elapsed:.1f} 秒")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    _cli()
