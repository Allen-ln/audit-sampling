# -*- coding: utf-8 -*-
"""
抽样底稿输出。

Sheet 编排：
  01_抽样汇总          总体、参数、样本量推导、结果指标、结论
  02_样本清单_PPS      （method 含 pps 时）
  03_样本清单_分层     （method 含 stratified 时）
  04_分层统计          （method 含 stratified 时）
  05_PPS选样轨迹       （method 含 pps 时）
  06_零负金额项目      红字冲销等，不进入抽样总体
  07_抽样方法备忘      可直接放进底稿的说明文字
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

from .sampler import SamplingRun

_TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="1F3864")
_HEAD_FONT = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
_HEAD_FILL = PatternFill("solid", fgColor="1F3864")
_BODY = Font(name="微软雅黑", size=10)
_BOLD = Font(name="微软雅黑", size=10, bold=True)
_MONO = Font(name="Consolas", size=9.5)
_GREEN = PatternFill("solid", fgColor="E2EFDA")
_YELLOW = PatternFill("solid", fgColor="FFF2CC")
_GREY = PatternFill("solid", fgColor="F2F2F2")
_THIN = Side(style="thin", color="D9D9D9")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_MONEY = "#,##0.00;[Red]-#,##0.00"


def _table(ws, df: pd.DataFrame, start_row: int, widths: dict[str, int] | None = None,
           money: tuple[str, ...] = (), pct: tuple[str, ...] = ()) -> int:
    if df is None or df.empty:
        ws.cell(row=start_row, column=1, value="（无记录）").font = _BODY
        return start_row + 2
    headers = list(df.columns)
    for j, name in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=j, value=name)
        c.font, c.fill, c.border = _HEAD_FONT, _HEAD_FILL, _BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(j)].width = (widths or {}).get(name, 14)
    for i, row in enumerate(df.to_dict("records"), start=start_row + 1):
        for j, name in enumerate(headers, start=1):
            v = row[name]
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime() if pd.notna(v) else None
            if isinstance(v, float) and pd.isna(v):
                v = None
            c = ws.cell(row=i, column=j, value=v)
            c.font, c.border = _BODY, _BORDER
            if name in money and isinstance(v, (int, float)):
                c.number_format = _MONEY
            elif name in pct and isinstance(v, (int, float)):
                c.number_format = "0.0%"
            elif name == "日期" and isinstance(v, datetime):
                c.number_format = "yyyy-mm-dd"
                c.alignment = Alignment(horizontal="center")
    return start_row + len(df) + 2


def _sample_frame(df: pd.DataFrame, method: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = pd.DataFrame({
        "序号": range(1, len(df) + 1),
        "原表行号": df["source_row"].astype(int),
        "凭证号": df["voucher"],
        "日期": df["date"],
        "摘要": df["summary"],
        "对方单位": df["counterparty"],
        "金额": df["amount"],
    })
    if method == "stratified" and "stratum" in df.columns:
        out["所属层"] = df["stratum"].astype(str)
    out["选中方式"] = df["selection"]
    if "hits" in df.columns:
        out["命中次数"] = df["hits"]
    out["选中概率"] = df["select_prob"]
    out["入选理由"] = df["reason"]
    return out


def _memo_text(run: SamplingRun) -> list[tuple[str, str]]:
    """生成底稿用的方法备忘（纯文字，符合底稿叙述习惯）。"""
    cfg = run.config
    info = run.sample_size_info
    pop = run.population
    lines: list[tuple[str, str]] = []

    lines.append(("一、抽样总体",
                  f"本次抽样总体为明细账记录，共 {len(pop)} 笔，账面金额合计 "
                  f"{run.book_value:,.2f} 元。另剔除金额为零或负数的记录 "
                  f"{len(run.excluded)} 笔（红字冲销等），单独列示于「06_零负金额项目」，"
                  f"未纳入抽样总体。"))

    tm = cfg.tolerable if cfg.tolerable is not None else cfg.materiality
    lines.append(("二、参数与样本量推导",
                  f"置信水平 {info['confidence_used']:.1f}%，对应可靠系数 "
                  f"{info['reliability_factor']:.2f}；可容忍错报 {tm:,.2f} 元，"
                  f"预期错报 {cfg.expected:,.2f} 元。"
                  f"抽样间隔 = （可容忍错报 - 预期错报）/ 可靠系数 = "
                  f"{info['tolerable_minus_expected']:,.2f} / {info['reliability_factor']:.2f} "
                  f"= {info['interval']:,.2f} 元；"
                  f"所需样本量 = 总体金额 / 抽样间隔 = "
                  f"{run.book_value:,.2f} / {info['interval']:,.2f} ≈ {info['required_size']} 笔。"))

    if run.pps is not None:
        p = run.pps
        lines.append(("三、PPS（货币单位）选样",
                      f"采用概率比例规模（PPS / MUS）选样。以随机种子 {cfg.seed} 在 "
                      f"[0, {p.interval:,.2f}) 区间内生成随机起点 {p.start_point:,.2f} 元，"
                      f"此后按 {p.interval:,.2f} 元间隔等距选取货币单位，共选取 "
                      f"{len(p.trace)} 个货币单位，涉及凭证 {p.distinct_items} 笔。"
                      f"金额达到或超过抽样间隔的项目必然被选中，已全部覆盖。"
                      f"选样轨迹见「05_PPS选样轨迹」，可逐点复核。"
                      f"需注意：PPS 下每个货币单位代表 {p.interval:,.2f} 元总体金额，"
                      f"因此样本项目的账面金额合计（{float(p.sample['amount'].sum()):,.2f} 元，"
                      f"占比 {float(p.sample['amount'].sum()) / run.book_value:.2%}）"
                      f"不等于货币单位覆盖率；错报推断采用错报比例法（tainting），"
                      f"即按样本项目的错报比例乘以抽样间隔推算总体错报。"))

    if run.strat is not None:
        st = run.strat
        lines.append(("四、分层选样",
                      f"按金额将总体分为 {len(st.stats)} 层，分层统计见「04_分层统计」。"
                      f"金额达到高值层阈值（{cfg.top_layer if cfg.top_layer else run.sample_size_info['interval']:,.2f} 元）"
                      f"的项目实施 100% 检查，共 {int(st.stats['拟抽样数'].iloc[-1])} 笔；"
                      f"其余各层按层内金额占比分配样本量，层内随机抽取，"
                      f"实际抽取 {st.actual_size} 笔。"))

    lines.append(("五、可复现性",
                  f"本底稿可通过相同参数与随机种子 {cfg.seed} 完全复现："
                  f"命令中指定 --seed {cfg.seed}、--confidence {cfg.confidence}、"
                  f"--tolerable {tm:,.2f}、--expected {cfg.expected:,.2f}，"
                  f"重新执行即得到完全相同的样本清单。"))

    lines.append(("六、使用限制",
                  "本工具仅完成选样环节。样本执行结果、错报汇总与推断（如存在错报需评估"
                  "是否扩大样本或调整重要性水平）仍需审计人员根据实际情况判断，"
                  "本表不构成审计结论。"))
    return lines


def write_report(run: SamplingRun, out_path: str | Path, meta: dict | None = None) -> Path:
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    meta = meta or {}
    cfg = run.config
    info = run.sample_size_info
    tm = cfg.tolerable if cfg.tolerable is not None else cfg.materiality

    wb = Workbook()
    ws = wb.active
    ws.title = "01_抽样汇总"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:D1")
    ws["A1"] = "审计抽样底稿 · 抽样结果汇总"
    ws["A1"].font = _TITLE_FONT
    ws.row_dimensions[1].height = 28
    ws.merge_cells("A2:D2")
    ws["A2"] = (f"生成时间：{datetime.now():%Y-%m-%d %H:%M}　"
                f"数据源：{Path(meta.get('data_file', '')).name}　"
                f"方法：{ {'pps': 'PPS 货币单位选样', 'stratified': '分层抽样', 'both': '两种方法并行'}[cfg.method] }")
    ws["A2"].font = Font(name="微软雅黑", size=9, color="808080")

    heads = ["项目", "数值", "单位", "说明"]
    for j, h in enumerate(heads, start=1):
        c = ws.cell(row=4, column=j, value=h)
        c.font, c.fill, c.border = _HEAD_FONT, _HEAD_FILL, _BORDER
        c.alignment = Alignment(horizontal="center")

    rows: list[tuple] = [
        ("总体项目数", len(run.population), "笔", f"剔除零负金额 {len(run.excluded)} 笔后"),
        ("总体账面金额", run.book_value, "元", "抽取样本所覆盖的总体金额"),
        ("置信水平", info["confidence_used"], "%", f"对应可靠系数 {info['reliability_factor']:.2f}"),
        ("可容忍错报", tm, "元", "TM"),
        ("预期错报", cfg.expected, "元", "EM"),
        ("抽样间隔", info["interval"], "元", "（TM - EM）/ 可靠系数"),
        ("所需样本量", info["required_size"], "笔", "总体金额 / 抽样间隔"),
    ]

    if run.pps is not None:
        p = run.pps
        sel_value = float(p.sample["amount"].sum())
        rows += [
            ("PPS 选中凭证数", p.distinct_items, "笔", f"共选取 {len(p.trace)} 个货币单位"),
            ("PPS 样本金额", round(sel_value, 2), "元", "被选中项目的账面金额合计"),
            ("PPS 金额覆盖率", round(sel_value / run.book_value * 100, 2), "%", "样本金额 / 总体金额"),
            ("PPS 随机起点", p.start_point, "元", f"随机种子 {cfg.seed}"),
        ]
    if run.strat is not None:
        st = run.strat
        sel_value = float(st.sample["amount"].sum()) if not st.sample.empty else 0.0
        rows += [
            ("分层样本数", st.actual_size, "笔", f"高值层全查 + 各层按比例"),
            ("分层样本金额", round(sel_value, 2), "元", ""),
            ("分层金额覆盖率", round(sel_value / run.book_value * 100, 2) if run.book_value else 0, "%", "样本金额 / 总体金额"),
        ]
    if cfg.method == "both" and not run.selected_any.empty:
        rows.append(("去重后涉及凭证数", int(run.selected_any["source_row"].nunique()), "笔",
                     "两种方法合计覆盖的项目数"))

    for i, (label, value, unit, note) in enumerate(rows, start=5):
        ws.cell(row=i, column=1, value=label).font = _BOLD
        v = ws.cell(row=i, column=2, value=value)
        v.font, v.alignment = _BOLD, Alignment(horizontal="right")
        if unit == "元":
            v.number_format = _MONEY
        elif unit == "%" and isinstance(value, (int, float)):
            v.number_format = "0.0" if value >= 10 else "0.00"
        ws.cell(row=i, column=3, value=unit).font = _BODY
        ws.cell(row=i, column=4, value=note).font = _BODY
        for j in range(1, 5):
            ws.cell(row=i, column=j).border = _BORDER
        if unit == "%" and label.endswith("覆盖率"):
            for j in range(1, 5):
                ws.cell(row=i, column=j).fill = _GREEN

    note_row = 5 + len(rows) + 1
    ws.cell(row=note_row, column=1, value="结论").font = _BOLD
    if cfg.method == "pps":
        actual_txt = (f"实际抽取 {run.pps.distinct_items} 笔（{len(run.pps.trace)} 个货币单位），"
                      f"覆盖总体账面金额的 {run.pps.sample['amount'].sum() / run.book_value * 100:.2f}%")
    elif cfg.method == "stratified":
        actual_txt = (f"实际抽取 {run.strat.actual_size} 笔，"
                      f"覆盖总体账面金额的 {run.strat.sample['amount'].sum() / run.book_value * 100:.2f}%")
    else:
        actual_txt = (f"PPS 抽取 {run.pps.distinct_items} 笔、分层抽取 {run.strat.actual_size} 笔，"
                      f"去重后合计涉及 {int(run.selected_any['source_row'].nunique())} 笔凭证")

    method_txt = {"pps": "PPS 货币单位选样", "stratified": "分层抽样", "both": "两种方法并行"}[cfg.method]
    concl = (
        f"本次采用{method_txt}，在 {info['confidence_used']:.0f}% 置信水平下，"
        f"按可容忍错报 {tm:,.2f} 元推导出抽样间隔 {info['interval']:,.2f} 元、"
        f"所需样本量 {info['required_size']} 笔；{actual_txt}。"
        f"样本清单见对应 sheet，选样过程可凭随机种子 {cfg.seed} 完全复现。"
    )
    ws.merge_cells(start_row=note_row + 1, start_column=1, end_row=note_row + 1, end_column=4)
    c = ws.cell(row=note_row + 1, column=1, value=concl)
    c.font, c.alignment = _BODY, Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[note_row + 1].height = 62

    ws.merge_cells(start_row=note_row + 3, start_column=1, end_row=note_row + 4, end_column=4)
    warn = ws.cell(row=note_row + 3, column=1,
                   value="⚠ 说明：本表为选样环节的辅助工具输出。样本执行、错报评估与审计结论"
                         "仍需由审计人员依据准则判断；若样本中发现错报，可能需要扩大样本或重新评估重要性水平。")
    warn.font = Font(name="微软雅黑", size=9, color="C00000")
    warn.alignment = Alignment(wrap_text=True, vertical="top")
    warn.fill = PatternFill("solid", fgColor="FCE4E4")

    for col, w in zip("ABCD", (24, 20, 8, 44)):
        ws.column_dimensions[col].width = w

    # ---------------- 样本清单
    widths = {"摘要": 34, "对方单位": 26, "凭证号": 16, "入选理由": 32, "所属层": 26, "日期": 12}
    if run.pps is not None:
        sheet = wb.create_sheet("02_样本清单_PPS")
        sheet.sheet_view.showGridLines = False
        df = _sample_frame(run.pps.sample, "pps")
        _table(sheet, df, 1, widths=widths, money=("金额",), pct=("选中概率",))
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(df.columns) if not df.empty else 1)}{max(len(df)+1,1)}"

    if run.strat is not None:
        sheet = wb.create_sheet("03_样本清单_分层")
        sheet.sheet_view.showGridLines = False
        df = _sample_frame(run.strat.sample, "stratified")
        _table(sheet, df, 1, widths=widths, money=("金额",), pct=("选中概率",))
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(df.columns) if not df.empty else 1)}{max(len(df)+1,1)}"

        sheet = wb.create_sheet("04_分层统计")
        sheet.sheet_view.showGridLines = False
        st = run.strat.stats.copy()
        if not run.strat.sample.empty:
            actual = run.strat.sample.groupby(run.strat.sample["stratum"].astype(str)).size()
            st["实际抽样数"] = st["层"].map(actual).fillna(0).astype(int)
            st["层内抽样比例"] = (st["实际抽样数"] / st["项目数"]).round(3)
        _table(sheet, st, 1, widths={"层": 34}, money=("层金额",), pct=("金额占比", "层内抽样比例"))
        sheet.freeze_panes = "A2"

    if run.pps is not None:
        sheet = wb.create_sheet("05_PPS选样轨迹")
        sheet.sheet_view.showGridLines = False
        _table(sheet, run.pps.trace, 1, widths={"命中凭证号": 18}, money=("货币单位位置", "命中金额"))
        sheet.freeze_panes = "A2"

    # ---------------- 零负金额
    sheet = wb.create_sheet("06_零负金额项目")
    sheet.sheet_view.showGridLines = False
    if run.excluded.empty:
        sheet.cell(row=1, column=1, value="总体中不存在零或负金额的记录。").font = _BODY
    else:
        ex = pd.DataFrame({
            "原表行号": run.excluded["source_row"].astype(int),
            "凭证号": run.excluded["voucher"],
            "日期": run.excluded["date"],
            "摘要": run.excluded["summary"],
            "对方单位": run.excluded["counterparty"],
            "金额": run.excluded["amount"],
        })
        _table(sheet, ex, 1, widths={"摘要": 34, "对方单位": 26}, money=("金额",))
        sheet.freeze_panes = "A2"

    # ---------------- 备忘
    sheet = wb.create_sheet("07_抽样方法备忘")
    sheet.sheet_view.showGridLines = False
    sheet.merge_cells("A1:B1")
    sheet["A1"] = "抽样方法备忘"
    sheet["A1"].font = _TITLE_FONT
    sheet.row_dimensions[1].height = 26
    r = 3
    for title, body in _memo_text(run):
        a = sheet.cell(row=r, column=1, value=title)
        a.font = Font(name="微软雅黑", size=10.5, bold=True, color="1F3864")
        b = sheet.cell(row=r + 1, column=1, value=body)
        b.font, b.alignment = _BODY, Alignment(wrap_text=True, vertical="top")
        sheet.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=2)
        # 行高按字数估算
        sheet.row_dimensions[r + 1].height = max(30, 16 * (len(body) // 46 + 1))
        r += 3
    sheet.column_dimensions["A"].width = 96
    sheet.column_dimensions["B"].width = 10

    for s in wb.worksheets:
        s.sheet_properties.tabColor = "1F3864"
        s.page_setup.orientation = "landscape"
        s.page_setup.fitToWidth = 1
        s.page_setup.fitToHeight = 0
        s.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

    wb.save(out)
    return out
