# -*- coding: utf-8 -*-
"""
抽样引擎：样本量推导 + PPS（MUS）选样 + 分层抽样选样。

设计原则与审计实务对齐：
1. 样本量不是拍脑袋定的，由置信水平、可容忍错报、预期错报推出来（MUS 公式）；
2. 选样必须可复现 —— 固定随机种子，同样的参数跑出同样的样本；
3. 高值项必须被覆盖 —— PPS 中金额超过抽样间隔的项目必然被选中；
4. 每个被选中的项目都要能解释"为什么选它"（金额比例 / 所在层 / 高值）。
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

# 置信水平 -> 可靠系数（MUS 常用表，来源：AICPA 审计抽样指南）
RELIABILITY_FACTORS = {
    99.0: 4.61,
    97.5: 3.84,
    95.0: 3.00,
    90.0: 2.31,
    85.0: 1.90,
    80.0: 1.61,
    75.0: 1.39,
    70.0: 1.22,
    60.0: 0.92,
    50.0: 0.70,
}


def reliability_factor(confidence: float) -> tuple[float, float]:
    """返回 (实际使用的置信水平, 可靠系数)，取表中最接近的档位。"""
    key = min(RELIABILITY_FACTORS, key=lambda k: abs(k - confidence))
    return key, RELIABILITY_FACTORS[key]


@dataclass
class SamplingConfig:
    method: str = "pps"            # pps | stratified | both
    confidence: float = 95.0
    materiality: float | None = None
    tolerable: float | None = None  # 可容忍错报 TM
    expected: float = 0.0           # 预期错报 EM
    seed: int = 20250101
    top_layer: float | None = None   # 高值层阈值，默认取抽样间隔
    strata_bounds: tuple[float, ...] | None = None
    min_per_stratum: int = 1


@dataclass
class PPSResult:
    sample: pd.DataFrame           # 被选中的项目（含 hits、选中概率）
    trace: pd.DataFrame            # 货币单位选样轨迹
    interval: float
    start_point: float
    required_size: int
    book_value: float
    distinct_items: int


@dataclass
class StratifiedResult:
    sample: pd.DataFrame
    stats: pd.DataFrame            # 分层统计
    required_size: int
    book_value: float
    top_layer_value: float
    actual_size: int


# ---------------------------------------------------------------- 样本量

def derive_sample_size(book_value: float, tolerable: float, expected: float,
                       confidence: float) -> dict:
    """按 MUS 公式推导抽样间隔与所需样本量。

    interval = (可容忍错报 - 预期错报) / 可靠系数
    样本量   = 总体金额 / 间隔
    """
    level, rf = reliability_factor(confidence)
    gap = tolerable - expected
    if gap <= 0:
        raise ValueError(
            f"可容忍错报（{tolerable:,.2f}）必须大于预期错报（{expected:,.2f}），"
            "否则抽多少样本都无法把风险降到可接受水平。"
        )
    interval = gap / rf
    size = math.ceil(book_value / interval) if interval > 0 else 0
    return {
        "confidence_used": level,
        "reliability_factor": rf,
        "tolerable_minus_expected": round(gap, 2),
        "interval": round(interval, 2),
        "required_size": size,
    }


# ---------------------------------------------------------------- PPS 选样

def pps_select(pop: pd.DataFrame, interval: float, seed: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """累计金额法 + 系统选样（MUS 的标准做法）。

    在 [0, interval) 内取一个随机起点，之后每隔 interval 取一个货币单位，
    货币单位落在哪个项目的累计金额区间，该项目就入选。
    """
    cum = pop["amount"].cumsum().to_numpy()
    total = float(cum[-1])
    rng = np.random.default_rng(seed)
    start = float(rng.uniform(0, interval))

    points = np.arange(start, total, interval)
    if points.size == 0:
        points = np.array([start])

    # 第一个累计金额 >= 选样点的位置
    idx = np.searchsorted(cum, points, side="left")
    idx = np.clip(idx, 0, len(pop) - 1)

    hits = pd.Series(idx).value_counts().sort_index()
    selected = pop.iloc[hits.index].copy()
    selected["hits"] = hits.values
    selected["selection"] = "PPS 货币单位选样"
    selected["select_prob"] = (selected["amount"] / interval).round(3)
    selected["reason"] = selected.apply(
        lambda r: ("金额超过抽样间隔，至少被选中一次" if r["amount"] >= interval
                   else f"货币单位命中 {int(r['hits'])} 次"), axis=1)
    selected["stratum"] = "—"

    trace = pd.DataFrame({
        "选样序号": range(1, len(points) + 1),
        "货币单位位置": np.round(points, 2),
        "命中凭证号": pop.iloc[idx]["voucher"].to_numpy(),
        "命中金额": pop.iloc[idx]["amount"].to_numpy(),
        "原表行号": pop.iloc[idx]["source_row"].to_numpy(),
    })

    return selected.sort_values("amount", ascending=False).reset_index(drop=True), trace


# ---------------------------------------------------------------- 分层抽样

def _auto_bounds(pop: pd.DataFrame, top_layer: float) -> tuple[float, ...]:
    """默认分层边界：按金额分 3 层（小 / 中 / 高值），高值层以抽样间隔为上界。"""
    amounts = pop["amount"]
    q1, q2 = amounts.quantile(0.5), amounts.quantile(0.9)
    bounds = sorted({0.0, round(float(q1), 2), round(float(q2), 2), round(float(top_layer), 2)})
    return tuple(bounds)


def stratified_select(pop: pd.DataFrame, config: SamplingConfig, required_size: int,
                      top_layer: float) -> StratifiedResult:
    """分层抽样：高值层全查，其余层按金额占比分配样本量，层内随机。"""
    bounds = config.strata_bounds or _auto_bounds(pop, top_layer)
    edges = list(bounds) + [float("inf")]

    pop = pop.copy()
    labels = []
    for i in range(len(edges) - 1):
        upper = "∞" if edges[i + 1] == float("inf") else f"{edges[i + 1]:,.0f}"
        labels.append(f"L{i+1} [{edges[i]:,.0f} ~ {upper})")
    pop["stratum"] = pd.cut(pop["amount"], bins=edges, right=False, labels=labels)

    stats_rows = []
    for name, grp in pop.groupby("stratum", observed=True):
        lower = float(grp["amount"].min())
        stats_rows.append({
            "层": str(name), "项目数": len(grp),
            "层金额": round(float(grp["amount"].sum()), 2),
            "金额占比": 0.0, "拟抽样数": 0,
        })
    stats = pd.DataFrame(stats_rows)
    total_value = float(pop["amount"].sum())
    stats["金额占比"] = (stats["层金额"] / total_value).round(4) if total_value else 0.0

    # 高值层 = 金额 >= top_layer
    high_mask = pop["amount"] >= top_layer
    high_count = int(high_mask.sum())
    high_value = float(pop.loc[high_mask, "amount"].sum())

    rng = np.random.default_rng(config.seed)
    chosen_rows: list[pd.DataFrame] = []

    # 1) 高值层全部检查
    for _, it in pop[high_mask].iterrows():
        chosen_rows.append(pd.DataFrame([{
            **it.to_dict(), "hits": 1, "selection": "高值层全查",
            "select_prob": 1.0, "reason": "金额达到或超过高值层阈值，100% 检查",
        }]))

    # 2) 其余层按金额占比分配剩余样本量
    low = pop[~high_mask]
    remaining = max(required_size - high_count, 0)
    if remaining == 0 and high_count > 0:
        remaining = 0
    low_value = float(low["amount"].sum())

    allocated = 0
    for name, grp in low.groupby("stratum", observed=True):
        share = (float(grp["amount"].sum()) / low_value) if low_value else 0.0
        want = int(round(remaining * share))
        want = max(want, config.min_per_stratum) if remaining > 0 else 0
        want = min(want, len(grp))
        if want == 0:
            continue
        picked = grp.sample(n=want, random_state=int(rng.integers(0, 2**31 - 1)))
        block = picked.copy()
        block["hits"] = 1
        block["selection"] = "层内随机"
        block["select_prob"] = round(want / len(grp), 3)
        block["reason"] = f"{name} 层内随机抽取 {want}/{len(grp)}"
        chosen_rows.append(block)
        allocated += want
        stats.loc[stats["层"] == str(name), "拟抽样数"] = want

    if not stats.empty:
        stats.loc[stats["层"] == stats["层"].iloc[-1], "拟抽样数"] = high_count

    sample = pd.concat(chosen_rows, ignore_index=True) if chosen_rows else pd.DataFrame()
    if not sample.empty:
        sample = sample.sort_values("amount", ascending=False).reset_index(drop=True)

    return StratifiedResult(
        sample=sample,
        stats=stats,
        required_size=required_size,
        book_value=total_value,
        top_layer_value=high_value,
        actual_size=len(sample),
    )


# ---------------------------------------------------------------- 统一入口

@dataclass
class SamplingRun:
    """一次抽样运行的完整结果。"""
    config: SamplingConfig
    population: pd.DataFrame
    excluded: pd.DataFrame          # 零 / 负金额项目
    book_value: float
    sample_size_info: dict
    pps: PPSResult | None = None
    strat: StratifiedResult | None = None
    selected_any: pd.DataFrame = field(default_factory=pd.DataFrame)


def run_sampling(pop_raw: pd.DataFrame, config: SamplingConfig) -> SamplingRun:
    """执行抽样。金额 <= 0 的项目单独列出（红字冲销需另行评估），不进入抽样总体。"""
    excluded = pop_raw[pop_raw["amount"] <= 0].copy()
    pop = pop_raw[pop_raw["amount"] > 0].copy().reset_index(drop=True)

    if pop.empty:
        raise ValueError("总体里没有金额大于 0 的项目，无法抽样。")

    book_value = float(pop["amount"].sum())

    # 未给可容忍错报时，按重要性水平的 100%（并提示用户确认）
    tolerable = config.tolerable if config.tolerable is not None else config.materiality
    if tolerable is None:
        raise ValueError("必须提供可容忍错报 --tolerable（或 --materiality）。")

    info = derive_sample_size(book_value, tolerable, config.expected, config.confidence)
    top_layer = config.top_layer if config.top_layer is not None else info["interval"]

    result = SamplingRun(
        config=config,
        population=pop,
        excluded=excluded,
        book_value=round(book_value, 2),
        sample_size_info=info,
    )

    if config.method in {"pps", "both"}:
        sample, trace = pps_select(pop, info["interval"], config.seed)
        result.pps = PPSResult(
            sample=sample, trace=trace,
            interval=info["interval"], start_point=float(trace["货币单位位置"].iloc[0]),
            required_size=info["required_size"], book_value=book_value,
            distinct_items=len(sample),
        )

    if config.method in {"stratified", "both"}:
        result.strat = stratified_select(pop, config, info["required_size"], top_layer)

    # 汇总被任一方法选中的凭证（both 模式下用于去重计数）
    frames = []
    if result.pps is not None:
        frames.append(result.pps.sample.assign(method="PPS"))
    if result.strat is not None:
        frames.append(result.strat.sample.assign(method="分层"))
    if frames:
        all_sel = pd.concat(frames, ignore_index=True)
        result.selected_any = all_sel
    return result
