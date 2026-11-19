# -*- coding: utf-8 -*-
"""审计抽样工具入口（分层抽样 / PPS 货币单位选样）。

示例：
    python run.py --data sample_data/应收账款明细.xlsx \
                  --method both \
                  --confidence 95 --tolerable 3000000 --expected 0 \
                  --out output/抽样结果.xlsx
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

from sampling.loader import load_population
from sampling.sampler import SamplingConfig, run_sampling
from sampling.report import write_report


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        description="审计抽样工具：按 MUS 公式推导样本量，支持 PPS 与分层抽样",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument("--data", required=True, help="总体明细（xlsx/xls/csv）")
    p.add_argument("--out", default="output/抽样结果.xlsx", help="输出底稿路径")
    p.add_argument("--method", choices=["pps", "stratified", "both"], default="both",
                   help="pps=货币单位选样；stratified=分层抽样；both=两套并行")
    p.add_argument("--confidence", type=float, default=95.0, help="置信水平（50-99）")
    p.add_argument("--materiality", type=float, default=None, help="整体重要性水平")
    p.add_argument("--tolerable", type=float, default=None,
                   help="可容忍错报 TM；未给时取 materiality")
    p.add_argument("--expected", type=float, default=0.0, help="预期错报 EM")
    p.add_argument("--seed", type=int, default=20250101, help="随机种子（决定样本，可复现）")
    p.add_argument("--top-layer", type=float, default=None, help="高值层阈值，默认取抽样间隔")
    p.add_argument("--strata", default=None,
                   help="分层边界，逗号分隔，如 0,50000,300000,1000000")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    t0 = time.perf_counter()
    print("① 读取总体 …")
    pop = load_population(args.data)
    positive = pop[pop["amount"] > 0]
    print(f"   {len(pop)} 笔记录，其中金额>0 的 {len(positive)} 笔，"
          f"合计 {positive['amount'].sum():,.2f} 元；零/负金额 {len(pop) - len(positive)} 笔另行列示。")

    bounds = None
    if args.strata:
        bounds = tuple(float(x) for x in args.strata.split(",") if x.strip())

    cfg = SamplingConfig(
        method=args.method,
        confidence=args.confidence,
        materiality=args.materiality,
        tolerable=args.tolerable,
        expected=args.expected,
        seed=args.seed,
        top_layer=args.top_layer,
        strata_bounds=bounds,
    )

    print("② 推导样本量 …")
    run = run_sampling(pop, cfg)
    info = run.sample_size_info
    print(f"   可靠系数 {info['reliability_factor']:.2f}，抽样间隔 {info['interval']:,.2f} 元，"
          f"所需样本量 {info['required_size']} 笔")

    print("③ 选样并输出底稿 …")
    out = write_report(run, args.out, meta={"data_file": args.data})
    elapsed = time.perf_counter() - t0

    print("\n" + "=" * 56)
    print("抽样完成")
    print("=" * 56)
    if run.pps is not None:
        sel_value = float(run.pps.sample["amount"].sum())
        print(f"PPS：选中 {run.pps.distinct_items} 笔 / 货币单位 {len(run.pps.trace)} 个，"
              f"样本金额 {sel_value:,.2f} 元，覆盖率 {sel_value / run.book_value:.2%}")
    if run.strat is not None:
        sel_value = float(run.strat.sample["amount"].sum()) if not run.strat.sample.empty else 0.0
        print(f"分层：抽取 {run.strat.actual_size} 笔，样本金额 {sel_value:,.2f} 元，"
              f"覆盖率 {(sel_value / run.book_value if run.book_value else 0):.2%}")
    if cfg.method == "both" and not run.selected_any.empty:
        print(f"两种方法合计涉及凭证 {run.selected_any['source_row'].nunique()} 笔（去重）")
    print(f"耗时 {elapsed:.2f} 秒　输出：{Path(out).resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
