# -*- coding: utf-8 -*-
"""明细数据读取与字段归一化。

抽样只需要四个字段：编号、日期、摘要、金额（其余列原样保留，供抽样后回溯）。
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

_ALIASES = {
    "voucher": ["凭证号", "凭证编号", "单据号", "凭证字号", "编号", "voucher", "no"],
    "date": ["记账日期", "凭证日期", "日期", "发生日期", "业务日期", "date"],
    "summary": ["摘要", "业务内容", "业务摘要", "说明", "summary", "description"],
    "counterparty": ["对方单位", "客户名称", "供应商", "对方", "单位名称", "counterparty"],
    "account": ["科目", "会计科目", "总账科目", "account"],
    "amount": ["借方金额", "金额", "发生额", "交易金额", "amount", "debit"],
}

_NOISE = re.compile(r"[^\d.\-]")


def _norm(s: str) -> str:
    t = str(s).strip().lower()
    t = re.sub(r"[（(].*?[)）]", "", t)
    return t.replace("　", "").replace(" ", "").replace(":", "")


def _find(cols: list[str], aliases: list[str]) -> str | None:
    normed = {_norm(c): c for c in cols}
    for a in aliases:
        if _norm(a) in normed:
            return normed[_norm(a)]
    for a in aliases:
        for key, original in normed.items():
            if _norm(a) and _norm(a) in key:
                return original
    return None


def clean_amount(value) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text in {"", "-", "—", "/", "nan", "None"}:
        return 0.0
    negative = text.startswith("(") and text.endswith(")")
    text = _NOISE.sub("", text)
    if text in {"", "-", ".", "-."}:
        return 0.0
    try:
        num = float(text)
    except ValueError:
        return 0.0
    return -num if negative else num


def _read(path: Path) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"找不到文件：{path}")
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path, dtype=str, sheet_name=0)
    if suffix in {".csv", ".txt"}:
        for enc in ("utf-8-sig", "gbk", "utf-8"):
            try:
                return pd.read_csv(path, dtype=str, encoding=enc)
            except UnicodeDecodeError:
                continue
        return pd.read_csv(path, dtype=str, encoding="utf-8", errors="replace")
    raise ValueError(f"不支持的文件类型：{suffix}")


def load_population(path: str | Path) -> pd.DataFrame:
    """读取总体明细。

    返回列：source_row（原表行号，便于回溯）、voucher、date、summary、
            counterparty、account、amount。
    """
    path = Path(path)
    raw = _read(path)
    cols = [str(c) for c in raw.columns]

    voucher_col = _find(cols, _ALIASES["voucher"])
    date_col = _find(cols, _ALIASES["date"])
    summary_col = _find(cols, _ALIASES["summary"])
    cp_col = _find(cols, _ALIASES["counterparty"])
    acct_col = _find(cols, _ALIASES["account"])
    amount_col = _find(cols, _ALIASES["amount"])

    if amount_col is None:
        raise ValueError(f"[{path.name}] 找不到金额列。当前列名：{cols}")
    if voucher_col is None and date_col is None:
        raise ValueError(f"[{path.name}] 找不到凭证号或日期列，无法保证抽样可回溯。当前列名：{cols}")

    out = pd.DataFrame()
    out["source_row"] = list(range(2, len(raw) + 2))
    out["voucher"] = raw[voucher_col].astype(str).str.strip() if voucher_col else \
        [f"ROW{i}" for i in out["source_row"]]
    if date_col:
        out["date"] = pd.to_datetime(
            raw[date_col].astype(str).str.strip().str.replace("/", "-", regex=False),
            errors="coerce")
    else:
        out["date"] = pd.NaT
    out["summary"] = raw[summary_col].astype(str).str.strip() if summary_col else ""
    out["counterparty"] = raw[cp_col].astype(str).str.strip() if cp_col else ""
    out["account"] = raw[acct_col].astype(str).str.strip() if acct_col else ""
    out["amount"] = raw[amount_col].map(clean_amount).round(2)

    out = out[~(out["amount"].isna() & out["voucher"].isna())].copy()
    return out.reset_index(drop=True)
