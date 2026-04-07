#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
经销商经营状况白皮书（HTML）生成器

目标：参考行业调研报告的表达方式，但用公司现有数据输出 15–20 个"分页段落"的经销商单页 HTML。

输入（默认）：
- xfx_inventory_tool/新家园进货销售*.xlsx（可用金额=库存）
- xfx_inventory_tool/经销商.xlsx（客户编码/主客户编码/名称）
- baipishu/data/经销商PTS计分卡*.xlsx（可选）
- baipishu/data/DSR分销能力*.xlsx（可选）

输出：
- jxsbps/output_html/{经销商名称}_{编码}_经销商白皮书.html
"""

from __future__ import annotations

import argparse
import base64
import io
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = Path(__file__).resolve().parent / "output_html"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def _safe_filename(name: str) -> str:
    s = re.sub(r"[\\/:*?\"<>|]", "_", str(name or "").strip())
    s = re.sub(r"\s+", " ", s).strip()
    return s[:120] or "未知经销商"


def _fmt_money(x: float | int | None) -> str:
    try:
        if x is None:
            return "—"
        v = float(x)
        return f"{v:,.0f}"
    except Exception:
        return "—"


def _pct(x: float | None) -> str:
    try:
        if x is None:
            return "—"
        return f"{float(x)*100:.1f}%"
    except Exception:
        return "—"


def _b64_png(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=160, bbox_inches="tight")
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("ascii")


def _load_inventory_detail() -> pd.DataFrame:
    # 复用 xfx_rules_core 的直连解析（可用金额=库存）
    import sys

    sys.path.insert(0, str(ROOT / "xfx_inventory_tool"))
    from xfx_rules_core import normalize_columns, prepare_direct_available_inventory

    # 与 copilot_inventory._inventory_path 一致：允许环境变量覆盖
    custom = os.environ.get("COPILOT_INVENTORY_XLSX", "").strip()
    if custom:
        path = Path(custom)
    else:
        path = ROOT / "xfx_inventory_tool" / "新家园进货销售刘杨20260323114325.xlsx"
    if not path.is_file():
        raise FileNotFoundError(f"未找到库存源文件：{path}")
    raw = pd.read_excel(path, sheet_name=0, engine="openpyxl")
    raw = normalize_columns(raw)
    detail = prepare_direct_available_inventory(raw, source_file=path.name)
    # 尽量补一个"经销商编码"列（如果原始表有）
    # 注：prepare_direct_available_inventory 会把经销商编码拼进 商品编码，故这里从 raw 再取一次更稳
    code_col = next((c for c in ["经销商编码", "客户编码", "编码"] if c in raw.columns), None)
    if code_col and code_col in raw.columns:
        detail["经销商编码"] = raw[code_col].astype(str).str.strip()
    else:
        detail["经销商编码"] = ""
    # 客户名称去空白
    detail["客户名称"] = detail["客户名称"].astype(str).str.strip()
    detail["商品名称"] = detail["商品名称"].astype(str).str.strip()
    detail["品类"] = detail["品类"].astype(str).str.strip()
    return detail


def _load_dealer_master() -> pd.DataFrame:
    path = ROOT / "xfx_inventory_tool" / "经销商.xlsx"
    if not path.is_file():
        return pd.DataFrame()
    df = pd.read_excel(path, sheet_name=0, engine="openpyxl").fillna("")
    # 常用字段：客户编码 / 主客户编码 / 客户名称 / 战区（若有）
    for c in ["客户编码", "主客户编码", "客户名称", "战区", "省", "市", "县", "城市", "客户类型"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df


def _pick_latest_xlsx(prefix: str, folder: Path) -> Path | None:
    if not folder.is_dir():
        return None
    cands = sorted(folder.glob(f"{prefix}*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return cands[0] if cands else None


def _load_pts() -> pd.DataFrame:
    # 优先从脚本同目录查找，兼容旧路径
    for search_dir in [Path(__file__).resolve().parent, ROOT / "baipishu" / "data"]:
        p = _pick_latest_xlsx("经销商PTS计分卡", search_dir)
        if p:
            return pd.read_excel(p, sheet_name=0, engine="openpyxl").fillna("")
    return pd.DataFrame()


def _load_dsr() -> pd.DataFrame:
    for search_dir in [Path(__file__).resolve().parent, ROOT / "baipishu" / "data"]:
        p = _pick_latest_xlsx("DSR分销能力", search_dir)
        if p:
            return pd.read_excel(p, sheet_name=0, engine="openpyxl").fillna("")
    return pd.DataFrame()


def _load_dealer_dist() -> pd.DataFrame:
    """经销商分销能力：经销商级别的网点/KOC/终端等级/品牌分销数据。"""
    for search_dir in [Path(__file__).resolve().parent, ROOT / "baipishu" / "data"]:
        p = _pick_latest_xlsx("经销商分销能力", search_dir)
        if p:
            df = pd.read_excel(p, sheet_name=0, engine="openpyxl").fillna(0)
            for c in ["经销商编码", "经销商名称", "大区", "区域", "城市", "客户经理工号", "客户经理名称", "经销商等级"]:
                if c in df.columns:
                    df[c] = df[c].astype(str).str.strip()
            return df
    return pd.DataFrame()


def _load_city_mgr() -> pd.DataFrame:
    """城市经理分销能力：城市经理维度的对标数据（用于区域背景展示）。"""
    for search_dir in [Path(__file__).resolve().parent, ROOT / "baipishu" / "data"]:
        p = _pick_latest_xlsx("城市经理分销能力", search_dir)
        if p:
            df = pd.read_excel(p, sheet_name=0, engine="openpyxl").fillna(0)
            for c in ["大区", "区域", "城市经理工号", "城市经理名称"]:
                if c in df.columns:
                    df[c] = df[c].astype(str).str.strip()
            return df
    return pd.DataFrame()


def _dealer_bucket_codes(dealer_df: pd.DataFrame, dealer_code: str) -> list[str]:
    """同一主客户簇下的所有客户编码（用于把主/关联户的数据合并到"该经销商报告"）。"""
    if dealer_df is None or dealer_df.empty:
        return [dealer_code]
    if "客户编码" not in dealer_df.columns or "主客户编码" not in dealer_df.columns:
        return [dealer_code]
    code = str(dealer_code).strip()
    if not code:
        return [dealer_code]
    hit = dealer_df[dealer_df["客户编码"].astype(str).str.strip().str.upper() == code.upper()]
    if hit.empty:
        # 尝试去掉 C 前缀
        code2 = code.upper().lstrip("C")
        hit = dealer_df[dealer_df["客户编码"].astype(str).str.strip().str.upper().str.lstrip("C") == code2]
    if hit.empty:
        return [dealer_code]
    main = str(hit.iloc[0].get("主客户编码", "")).strip() or code
    main_u = main.upper()
    # bucket: 主客户编码=main 的所有客户编码 + 主客户自身
    bucket = dealer_df[
        dealer_df["主客户编码"].astype(str).str.strip().str.upper().str.lstrip("C") == main_u.lstrip("C")
    ]
    codes = (
        bucket["客户编码"].astype(str).str.strip().tolist()
        if "客户编码" in bucket.columns
        else [code]
    )
    codes = [c for c in codes if str(c).strip()]
    if code not in codes:
        codes.append(code)
    # 去重
    out = []
    seen = set()
    for c in codes:
        k = str(c).strip().upper()
        if k not in seen:
            seen.add(k)
            out.append(str(c).strip())
    return out or [dealer_code]


def _rank_pct(series: pd.Series, value: float) -> float | None:
    try:
        s = pd.to_numeric(series, errors="coerce").dropna()
        if s.empty:
            return None
        # 越大越好：百分位（0-1）
        return float((s <= value).mean())
    except Exception:
        return None


def build_dealer_report_html(
    *,
    dealer_code: str,
    dealer_name: str,
    bucket_codes: list[str],
    inv_detail: pd.DataFrame,
    inv_all_agg: pd.DataFrame,
    pts_df: pd.DataFrame,
    dsr_df: pd.DataFrame,
    dealer_dist_df: pd.DataFrame,
    city_mgr_df: pd.DataFrame,
    dealer_master: pd.DataFrame,
) -> str:
    # 过滤明细（按经销商编码"簇"）
    dc_set = {str(x).strip().upper() for x in bucket_codes if str(x).strip()}
    sub = inv_detail.copy()
    if "经销商编码" in sub.columns and dc_set:
        m = sub["经销商编码"].astype(str).str.strip().str.upper().isin(dc_set)
        sub = sub[m]
    # 若编码缺失，则按名称兜底
    if sub.empty and dealer_name:
        m2 = sub["客户名称"].astype(str).str.strip() == str(dealer_name).strip()
        sub = sub[m2]

    # 基础 KPI
    inv_sum = float(sub["库存额"].sum()) if not sub.empty else 0.0
    in_sum = float(sub["进货额"].sum()) if not sub.empty else 0.0
    sale_sum = float(sub["销售额出厂价"].sum()) if not sub.empty else 0.0
    neg_rows = int((sub["库存额"] < 0).sum()) if (not sub.empty and "库存额" in sub.columns) else 0
    sku_rows = int(len(sub))
    cat_cnt = int(sub["品类"].nunique()) if (not sub.empty and "品类" in sub.columns) else 0

    # 库销比（避免除 0）
    turnover_ratio = None
    if sale_sum > 0:
        turnover_ratio = inv_sum / sale_sum

    # 全体对标分位
    pct_inv = _rank_pct(inv_all_agg["库存额"], inv_sum) if not inv_all_agg.empty else None
    pct_sale = _rank_pct(inv_all_agg["销售额出厂价"], sale_sum) if not inv_all_agg.empty else None

    # 品类结构
    cat = (
        sub.groupby("品类", as_index=False)[["库存额", "进货额", "销售额出厂价"]].sum()
        if not sub.empty
        else pd.DataFrame(columns=["品类", "库存额", "进货额", "销售额出厂价"])
    )
    cat = cat.sort_values("库存额", ascending=False, key=lambda s: s.abs()).head(10)

    # Top SKU（按库存额绝对值）
    top_sku = (
        sub.sort_values("库存额", ascending=False, key=lambda s: s.abs()).head(15)
        if not sub.empty
        else pd.DataFrame()
    )

    cat_rows_html = ""
    if cat is not None and not cat.empty:
        rows = []
        for c, j, x, k in cat[["品类", "进货额", "销售额出厂价", "库存额"]].itertuples(
            index=False, name=None
        ):
            rows.append(
                "<tr>"
                f"<td>{c}</td>"
                f"<td>{_fmt_money(j)}</td>"
                f"<td>{_fmt_money(x)}</td>"
                f"<td>{_fmt_money(k)}</td>"
                "</tr>"
            )
        cat_rows_html = "".join(rows)
    else:
        cat_rows_html = '<tr><td colspan="4" class="muted">无数据</td></tr>'

    top_sku_rows_html = ""
    if top_sku is not None and not top_sku.empty:
        rows = []
        for i, (n, p, j, x, k, st) in enumerate(
            top_sku[
                ["商品名称", "品类", "进货额", "销售额出厂价", "库存额", "状态"]
            ].itertuples(index=False, name=None)
        ):
            color = "#ef4444" if float(k) < 0 else "#111827"
            rows.append(
                "<tr>"
                f"<td>{i+1}</td>"
                f"<td>{n}</td>"
                f"<td>{p}</td>"
                f"<td>{_fmt_money(j)}</td>"
                f"<td>{_fmt_money(x)}</td>"
                f"<td><b style='color:{color}'>{_fmt_money(k)}</b></td>"
                f"<td>{st}</td>"
                "</tr>"
            )
        top_sku_rows_html = "".join(rows)
    else:
        top_sku_rows_html = '<tr><td colspan="7" class="muted">无数据</td></tr>'

    # ── PTS 计分卡 ──
    pts_info = {}
    if not pts_df.empty:
        cand_cols = [c for c in ["经销商编码", "客户编码", "客户编号", "经销商编号"] if c in pts_df.columns]
        hit = pd.DataFrame()
        for c in cand_cols:
            s = pts_df[c].astype(str).str.strip().str.upper()
            if dealer_code and (s == dealer_code.upper()).any():
                hit = pts_df[s == dealer_code.upper()]
                break
        if hit.empty and dealer_name and "经销商名称" in pts_df.columns:
            hit = pts_df[pts_df["经销商名称"].astype(str).str.strip() == dealer_name]
        if not hit.empty:
            r = hit.iloc[0].to_dict()
            for k in ["客户等级", "当月得分", "季度得分", "季度排名", "年度得分", "年度排名",
                      "大区", "区域", "城市", "客户经理编码", "客户经理姓名", "主客户编码", "主客户名称"]:
                if k in r and r[k] not in ("", None, 0, "0"):
                    pts_info[k] = r[k]

    # ── DSR 分销能力（按经销商编码聚合多个DSR行）──
    dsr_info: dict = {}
    dsr_rows: pd.DataFrame = pd.DataFrame()
    if not dsr_df.empty:
        code_cols = [c for c in ["经销商编码", "客户编码"] if c in dsr_df.columns]
        for c in code_cols:
            s = dsr_df[c].astype(str).str.strip().str.upper()
            if dealer_code and (s == dealer_code.upper()).any():
                dsr_rows = dsr_df[s == dealer_code.upper()].copy()
                break
        if not dsr_rows.empty:
            num_cols = ["整体网点数", "活跃网点", "KOC网点数", "KOC销售额",
                        "拜访门店数", "拜访成交门店数", "分销门店", "销售额"]
            for col in num_cols:
                if col in dsr_rows.columns:
                    dsr_info[col] = float(pd.to_numeric(dsr_rows[col], errors="coerce").fillna(0).sum())
            # 拜访成交率：加权平均
            if "拜访成交率" in dsr_rows.columns:
                v = pd.to_numeric(dsr_rows["拜访成交率"], errors="coerce").mean()
                dsr_info["拜访成交率"] = float(v) if pd.notna(v) else None
            if "覆盖率" in dsr_rows.columns:
                v = pd.to_numeric(dsr_rows["覆盖率"], errors="coerce").sum()
                dsr_info["覆盖率"] = float(v) if pd.notna(v) else None

    # ── 经销商分销能力（dealer_dist_df）──
    dist_info: dict = {}
    if not dealer_dist_df.empty:
        code_cols = [c for c in ["经销商编码"] if c in dealer_dist_df.columns]
        hit = pd.DataFrame()
        for c in code_cols:
            s = dealer_dist_df[c].astype(str).str.strip().str.upper()
            if dealer_code and (s == dealer_code.upper()).any():
                hit = dealer_dist_df[s == dealer_code.upper()]
                break
        if hit.empty and dealer_name and "经销商名称" in dealer_dist_df.columns:
            hit = dealer_dist_df[dealer_dist_df["经销商名称"].astype(str).str.strip() == dealer_name]
        if not hit.empty:
            r = hit.iloc[0]
            scalar_fields = [
                "经销商等级", "整体网点数", "KOC网点数", "KOC销售额",
                "市场容量", "覆盖率(%)", "活跃网点", "品牌分销指数", "复用指数", "经销商分销能力",
                "本月金标店数", "本月银标店数", "本月铜标店数", "本月基础店数", "本月不达标数",
                "南孚网点数", "丰蓝网点数", "益圆网点数", "传应网点数",
                "KOC-火机品类网点数", "KOC-辣味零食品类网点数", "KOC-爆珠网点数", "KOC-剃须刀品类网点数",
                "大区", "区域", "城市", "客户经理名称",
            ]
            for f in scalar_fields:
                if f in r.index and r[f] not in ("", None):
                    dist_info[f] = r[f]

    # ── 城市经理对标（取该经销商所在区域的城市经理行，用于背景展示）──
    cm_info: dict = {}
    if not city_mgr_df.empty and (dist_info.get("区域") or pts_info.get("区域")):
        target_region = str(dist_info.get("区域") or pts_info.get("区域", "")).strip()
        if target_region and "区域" in city_mgr_df.columns:
            cm_rows = city_mgr_df[city_mgr_df["区域"].astype(str).str.strip() == target_region]
            if not cm_rows.empty:
                cm_r = cm_rows.iloc[0]
                for f in ["城市经理名称", "整体网点数", "KOC网点数", "KOC销售额", "覆盖率(%)", "分销能力", "市场容量"]:
                    if f in cm_r.index:
                        cm_info[f] = cm_r[f]

    # ── Matplotlib 图表 ──
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({"font.sans-serif": ["SimHei", "Microsoft YaHei", "STHeiti"], "axes.unicode_minus": False})

    # 图1：品类库存 TOP10
    fig1 = plt.figure(figsize=(7.2, 3.2))
    ax = fig1.add_subplot(111)
    if not cat.empty:
        ax.barh(cat["品类"].astype(str).iloc[::-1], cat["库存额"].astype(float).iloc[::-1], color="#1a3a5c")
        ax.set_title("品类库存 TOP10（库存额=可用金额）", fontsize=11)
        ax.tick_params(labelsize=9)
    else:
        ax.text(0.5, 0.5, "无可用品类数据", ha="center", va="center")
        ax.axis("off")
    fig1.tight_layout()
    img_cat = _b64_png(fig1)
    plt.close(fig1)

    # 图2：进销存概览
    fig2 = plt.figure(figsize=(7.2, 2.8))
    ax2 = fig2.add_subplot(111)
    ax2.bar(["进货额", "销售额(出厂)", "库存额"], [in_sum, sale_sum, inv_sum],
            color=["#3b82f6", "#10b981", "#1a3a5c"])
    ax2.set_title("进销存概览（当期汇总）", fontsize=11)
    ax2.tick_params(labelsize=9)
    fig2.tight_layout()
    img_kpi = _b64_png(fig2)
    plt.close(fig2)

    # 图3：终端等级分布（来自经销商分销能力）
    img_grade = ""
    grade_vals = {
        "金标": float(dist_info.get("本月金标店数", 0) or 0),
        "银标": float(dist_info.get("本月银标店数", 0) or 0),
        "铜标": float(dist_info.get("本月铜标店数", 0) or 0),
        "基础": float(dist_info.get("本月基础店数", 0) or 0),
        "不达标": float(dist_info.get("本月不达标数", 0) or 0),
    }
    total_grade = sum(grade_vals.values())
    if total_grade > 0:
        fig3, ax3 = plt.subplots(figsize=(5.5, 3.0))
        colors = ["#f59e0b", "#94a3b8", "#b45309", "#64748b", "#ef4444"]
        bars = ax3.bar(list(grade_vals.keys()), list(grade_vals.values()), color=colors)
        ax3.set_title("终端门店等级分布", fontsize=11)
        ax3.tick_params(labelsize=9)
        for bar, v in zip(bars, grade_vals.values()):
            if v > 0:
                ax3.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                         f"{int(v)}", ha="center", va="bottom", fontsize=9)
        fig3.tight_layout()
        img_grade = _b64_png(fig3)
        plt.close(fig3)

    # 图4：品类 KOC 网点分布
    img_koc = ""
    koc_brand_map = {
        "火机": float(dist_info.get("KOC-火机品类网点数", 0) or 0),
        "辣味零食": float(dist_info.get("KOC-辣味零食品类网点数", 0) or 0),
        "爆珠": float(dist_info.get("KOC-爆珠网点数", 0) or 0),
        "剃须刀": float(dist_info.get("KOC-剃须刀品类网点数", 0) or 0),
    }
    koc_brand_map = {k: v for k, v in koc_brand_map.items() if v > 0}
    if koc_brand_map:
        fig4, ax4 = plt.subplots(figsize=(5.5, 3.0))
        ax4.barh(list(koc_brand_map.keys()), list(koc_brand_map.values()), color="#0ea5e9")
        ax4.set_title("品类 KOC 网点数", fontsize=11)
        ax4.tick_params(labelsize=9)
        fig4.tight_layout()
        img_koc = _b64_png(fig4)
        plt.close(fig4)

    # 报告日期
    report_date = datetime.now().strftime("%Y-%m-%d")

    # 15-20 个分页段落（打印时每段一页）
    pages = []

    _dealer_grade = dist_info.get("经销商等级") or pts_info.get("客户等级") or "—"
    _region = dist_info.get("区域") or pts_info.get("区域") or "—"
    _city = dist_info.get("城市") or pts_info.get("城市") or ""
    _km = dist_info.get("客户经理名称") or pts_info.get("客户经理姓名") or "—"

    pages.append(
        f"""
        <section class="page cover">
          <div class="cover-top">经销商经营状况白皮书 &nbsp;·&nbsp; 2025</div>
          <div class="cover-title">{dealer_name or "—"}</div>
          <div class="cover-sub">
            {_region}{" · " + _city if _city and _city != "—" else ""} &nbsp;|&nbsp;
            编码：{dealer_code} &nbsp;|&nbsp; 等级：{_dealer_grade} &nbsp;|&nbsp; 客户经理：{_km}
          </div>
          <div class="cover-meta">生成日期：{report_date} &nbsp;·&nbsp; 数据来源：进货销售表 / 经销商分销能力 / PTS计分卡 / DSR分销能力</div>
          <div class="cover-badges">
            <span class="badge">库存额 &nbsp;{_fmt_money(inv_sum)}</span>
            <span class="badge">销售额 &nbsp;{_fmt_money(sale_sum)}</span>
            <span class="badge">整体网点 &nbsp;{_fmt_money(dist_info.get("整体网点数") or dsr_info.get("整体网点数"))}</span>
            <span class="badge">KOC网点 &nbsp;{_fmt_money(dist_info.get("KOC网点数") or dsr_info.get("KOC网点数"))}</span>
          </div>
        </section>
        """
    )

    _pts_monthly = pts_info.get("当月得分", "—")
    _pts_quarter = pts_info.get("季度得分", "—")
    _pts_q_rank = pts_info.get("季度排名", "—")
    _pts_annual = pts_info.get("年度得分", "—")
    _pts_a_rank = pts_info.get("年度排名", "—")
    _dist_ability = dist_info.get("经销商分销能力", "—")
    _coverage = dist_info.get("覆盖率(%)", "—")
    _coverage_str = f"{float(_coverage):.1f}%" if _coverage not in ("—", None, 0, "") else "—"

    pages.append(
        f"""
        <section class="page">
          <h2>01 · 结论摘要（Executive Summary）</h2>
          <div class="grid3">
            <div class="card"><div class="k">库存合计</div><div class="v">{_fmt_money(inv_sum)}</div><div class="s">全体分位：{_pct(pct_inv)}</div></div>
            <div class="card"><div class="k">销售额（出厂）</div><div class="v">{_fmt_money(sale_sum)}</div><div class="s">全体分位：{_pct(pct_sale)}</div></div>
            <div class="card"><div class="k">库销比</div><div class="v">{(f"{turnover_ratio:.1f}x" if turnover_ratio is not None else "—")}</div><div class="s">销售=0 时不计算</div></div>
          </div>
          <div class="grid3">
            <div class="card"><div class="k">PTS 季度得分</div><div class="v">{_pts_quarter}</div><div class="s">季度排名 #{_pts_q_rank} &nbsp;|&nbsp; 当月 {_pts_monthly}</div></div>
            <div class="card"><div class="k">分销能力指数</div><div class="v">{_dist_ability}</div><div class="s">市场覆盖率 {_coverage_str}</div></div>
            <div class="card"><div class="k">整体网点数</div><div class="v">{_fmt_money(dist_info.get("整体网点数") or dsr_info.get("整体网点数"))}</div><div class="s">活跃网点 {_fmt_money(dist_info.get("活跃网点") or dsr_info.get("活跃网点"))}</div></div>
          </div>
          <div class="note">
            <b>解读建议：</b>库销比高 → 积压风险优先去化；分销能力低 → 聚焦网点质量与 KOC 渗透；PTS 排名滞后 → 核对拜访得分与规模系数短板。
          </div>
        </section>
        """
    )

    pages.append(
        f"""
        <section class="page">
          <h2>02 · 核心指标总览</h2>
          <img class="img" src="data:image/png;base64,{img_kpi}" />
          <div class="kpi-row">
            <div class="pill">系列行数：<b>{sku_rows}</b></div>
            <div class="pill">品类数：<b>{cat_cnt}</b></div>
            <div class="pill">负库存行：<b style="color:#ef4444">{neg_rows}</b></div>
          </div>
          <div class="small muted">说明：本报告不包含利润/费用/应收等财务指标；如需对齐行业调研报告的财务篇章，需要追加财务数据源。</div>
        </section>
        """
    )

    pages.append(
        f"""
        <section class="page">
          <h2>03 · 品类结构与重点品类</h2>
          <img class="img" src="data:image/png;base64,{img_cat}" />
          <table class="tbl">
            <thead><tr><th>品类</th><th>进货额</th><th>销售额</th><th>库存额</th></tr></thead>
            <tbody>
              {cat_rows_html}
            </tbody>
          </table>
        </section>
        """
    )

    pages.append(
        f"""
        <section class="page">
          <h2>04 · 系列明细 TOP（用于定位积压）</h2>
          <table class="tbl">
            <thead><tr><th>#</th><th>系列/商品名称</th><th>品类</th><th>进货额</th><th>销售额</th><th>库存额</th><th>状态</th></tr></thead>
            <tbody>
              {top_sku_rows_html}
            </tbody>
          </table>
          <div class="small muted">提示：优先关注"库存额高且销售低"的系列；若出现大量负库存，请先核对规则起算日/重复进表/冲减口径。</div>
        </section>
        """
    )

    # 风险诊断（规则化）
    risk_lines = []
    if sale_sum <= 0 and inv_sum > 0:
        risk_lines.append("停滞风险：有库存但销售额为 0（可能为未动销或口径缺失）。")
    if turnover_ratio is not None and turnover_ratio >= 3:
        risk_lines.append(f"库销比偏高：{turnover_ratio:.1f}x（建议做去化与进货闸门）。")
    if neg_rows > 0:
        risk_lines.append(f"负库存异常：共 {neg_rows} 行（建议先核对数据口径/时点）。")
    if not risk_lines:
        risk_lines = ["暂无明显规则风险信号（仍建议按品类与系列做结构性复盘）。"]

    pages.append(
        f"""
        <section class="page">
          <h2>05 · 经营风险雷达（规则诊断）</h2>
          <div class="card">
            <div class="k">自动诊断结论</div>
            <ul class="ul">{''.join([f'<li>{x}</li>' for x in risk_lines])}</ul>
          </div>
          <div class="grid2">
            <div class="card">
              <div class="k">动作建议（本周）</div>
              <ol class="ol">
                <li>按库存额排序锁定 TOP3 品类与 TOP10 系列；对"库存高/销售低"做去化动作。</li>
                <li>对库销比高的品类设置进货闸门，优先消化库存再补货。</li>
                <li>若负库存行多：先核"起算日/重复进表/冲减"，再谈补货或去化。</li>
              </ol>
            </div>
            <div class="card">
              <div class="k">动作建议（本月）</div>
              <ol class="ol">
                <li>建立"品类-系列"周度复盘：金额、动销、异常行、责任人、截止日期。</li>
                <li>对重点品类做陈列/促销资源申请并跟踪验收。</li>
                <li>把库存与动销指标纳入经销商月度例会，形成闭环。</li>
              </ol>
            </div>
          </div>
        </section>
        """
    )

    # 页06：经销商分销能力（核心新页）
    _active_rate = ""
    _total_net = float(dist_info.get("整体网点数") or 0)
    _active_net = float(dist_info.get("活跃网点") or 0)
    if _total_net > 0 and _active_net > 0:
        _active_rate = f"活跃率 {_active_net/_total_net*100:.0f}%"
    _brand_idx = dist_info.get("品牌分销指数", "—")
    _reuse_idx = dist_info.get("复用指数", "—")
    _dist_cap = dist_info.get("经销商分销能力", "—")
    _mkt_cap = dist_info.get("市场容量", "—")
    _koc_sales = dist_info.get("KOC销售额") or dsr_info.get("KOC销售额")
    pages.append(
        f"""
        <section class="page">
          <h2>06 · 分销覆盖与网点能力</h2>
          <div class="grid3">
            <div class="card"><div class="k">整体网点数</div><div class="v">{_fmt_money(_total_net or dist_info.get("整体网点数") or dsr_info.get("整体网点数"))}</div><div class="s">{_active_rate}</div></div>
            <div class="card"><div class="k">KOC 网点数</div><div class="v">{_fmt_money(dist_info.get("KOC网点数") or dsr_info.get("KOC网点数"))}</div><div class="s">KOC 销售额 {_fmt_money(_koc_sales)}</div></div>
            <div class="card"><div class="k">市场覆盖率</div><div class="v">{_coverage_str}</div><div class="s">市场容量 {_fmt_money(_mkt_cap)}</div></div>
          </div>
          <div class="grid3">
            <div class="card"><div class="k">品牌分销指数</div><div class="v">{_brand_idx}</div><div class="s">越高品牌铺货越广</div></div>
            <div class="card"><div class="k">复用指数</div><div class="v">{_reuse_idx}</div><div class="s">跨品类复用门店能力</div></div>
            <div class="card"><div class="k">经销商分销能力</div><div class="v">{_dist_cap}</div><div class="s">综合分销能力评分</div></div>
          </div>
          {"<img class='img' src='data:image/png;base64," + img_koc + "'/>" if img_koc else ""}
          <div class="small muted">说明：网点数来自经销商分销能力表；覆盖率 = 整体网点数 / 市场容量 × 100%。</div>
        </section>
        """
    )

    # 页07：终端等级分布
    _grade_html = ""
    if total_grade > 0:
        pcts = {k: f"{v/total_grade*100:.0f}%" for k, v in grade_vals.items()}
        rows = "".join(
            f"<tr><td>{k}</td><td>{int(v)}</td><td>{pcts[k]}</td></tr>"
            for k, v in grade_vals.items()
        )
        _grade_html = f"""
          <table class="tbl" style="max-width:360px">
            <thead><tr><th>等级</th><th>门店数</th><th>占比</th></tr></thead>
            <tbody>{rows}<tr><td><b>合计</b></td><td><b>{int(total_grade)}</b></td><td>100%</td></tr></tbody>
          </table>"""

    pages.append(
        f"""
        <section class="page">
          <h2>07 · 终端门店质量分布</h2>
          {"<img class='img' src='data:image/png;base64," + img_grade + "'/>" if img_grade else "<div class='note'>终端等级数据缺失（经销商分销能力表无对应编码）。</div>"}
          {_grade_html}
          <div class="small muted" style="margin-top:12px">金标 → 银标 → 铜标 → 基础 → 不达标；提升路径：不达标 → 基础 → 铜标优先，KOC 门店着重保金/银。</div>
        </section>
        """
    )

    # 页08：PTS 经营评分
    pages.append(
        f"""
        <section class="page">
          <h2>08 · PTS 经营评分与等级</h2>
          <div class="grid3">
            <div class="card"><div class="k">客户等级</div><div class="v">{pts_info.get("客户等级","—")}</div><div class="s">经销商分级</div></div>
            <div class="card"><div class="k">当月得分</div><div class="v">{_pts_monthly}</div><div class="s">最新月度评分</div></div>
            <div class="card"><div class="k">季度得分</div><div class="v">{_pts_quarter}</div><div class="s">季度排名 #{_pts_q_rank}</div></div>
          </div>
          <div class="grid3">
            <div class="card"><div class="k">年度得分</div><div class="v">{_pts_annual}</div><div class="s">年度排名 #{_pts_a_rank}</div></div>
            <div class="card"><div class="k">客户经理</div><div class="v" style="font-size:18px">{pts_info.get("客户经理姓名","—")}</div><div class="s">编码 {pts_info.get("客户经理编码","—")}</div></div>
            <div class="card"><div class="k">所属区域</div><div class="v" style="font-size:18px">{pts_info.get("区域","—")}</div><div class="s">{pts_info.get("城市","—")}</div></div>
          </div>
          <div class="small muted">说明：PTS 分数越高表示经营表现越好；季度/年度排名越小越靠前。数据来自经销商PTS计分卡。</div>
        </section>
        """
    )

    # 页09：DSR 分销员执行力
    _dsr_visit_rate = ""
    if dsr_info.get("拜访成交率"):
        _dsr_visit_rate = f"{float(dsr_info['拜访成交率'])*100:.1f}%" if float(dsr_info['拜访成交率']) <= 1 else f"{float(dsr_info['拜访成交率']):.1f}%"
    _dsr_rows_html = ""
    if not dsr_rows.empty:
        show_cols = [c for c in ["DSR名称", "整体网点数", "活跃网点", "KOC网点数", "拜访门店数", "拜访成交门店数", "销售额"] if c in dsr_rows.columns]
        rows_html = []
        for _, row in dsr_rows[show_cols].iterrows():
            cells = "".join(f"<td>{row[c]}</td>" for c in show_cols)
            rows_html.append(f"<tr>{cells}</tr>")
        headers = "".join(f"<th>{c}</th>" for c in show_cols)
        _dsr_rows_html = f"""
          <table class="tbl">
            <thead><tr>{headers}</tr></thead>
            <tbody>{"".join(rows_html)}</tbody>
          </table>"""
    pages.append(
        f"""
        <section class="page">
          <h2>09 · DSR 分销员执行力明细</h2>
          <div class="grid3">
            <div class="card"><div class="k">拜访门店数</div><div class="v">{_fmt_money(dsr_info.get("拜访门店数"))}</div></div>
            <div class="card"><div class="k">拜访成交门店</div><div class="v">{_fmt_money(dsr_info.get("拜访成交门店数"))}</div></div>
            <div class="card"><div class="k">拜访成交率</div><div class="v">{_dsr_visit_rate or "—"}</div></div>
          </div>
          {_dsr_rows_html if _dsr_rows_html else '<div class="note">DSR 明细数据暂无（编码未对齐）。</div>'}
          <div class="small muted" style="margin-top:12px">拜访成交率 = 拜访成交门店数 / 拜访门店数；目标 > 60%。</div>
        </section>
        """
    )

    # 页10：区域城市经理对标
    _cm_name = cm_info.get("城市经理名称", "—")
    _cm_region = dist_info.get("区域") or pts_info.get("区域", "—")
    pages.append(
        f"""
        <section class="page">
          <h2>10 · 区域背景与城市经理对标</h2>
          <div class="grid3">
            <div class="card"><div class="k">所属区域</div><div class="v" style="font-size:18px">{_cm_region}</div><div class="s">城市经理：{_cm_name}</div></div>
            <div class="card"><div class="k">区域整体网点</div><div class="v">{_fmt_money(cm_info.get("整体网点数"))}</div><div class="s">区域 KOC 网点 {_fmt_money(cm_info.get("KOC网点数"))}</div></div>
            <div class="card"><div class="k">区域覆盖率</div><div class="v">{str(cm_info.get("覆盖率(%)","—")).rstrip("0").rstrip(".") + "%" if cm_info.get("覆盖率(%)") else "—"}</div><div class="s">区域分销能力 {cm_info.get("分销能力","—")}</div></div>
          </div>
          <div class="note">
            <b>对标方向：</b>对比本经销商与区域城市经理汇总数据，识别该经销商在区域中的贡献占比与能力水位。覆盖率/分销能力明显低于区域均值时，优先做网点开发与 KOC 渗透。
          </div>
        </section>
        """
    )

    # 11-18：行动手册
    action_pages = [
        ("11 · 本周必做清单", ["锁定 TOP3 品类与 TOP10 系列", "对停滞/高库销比系列制定去化动作", "负库存先核口径再下结论", "核对终端不达标门店并制定提升方案"]),
        ("12 · 进货闸门与补货策略", ["积压品类：先去化再补货", "畅销品类：补货小步快跑", "新品：试销-复盘-再铺货", "KOC 品类优先保障货源"]),
        ("13 · 终端动销抓手（可验收）", ["陈列位提升（拍照验收）", "促销资源申请（ROI 复盘）", "KOC 门店渗透（名单+节奏）", "不达标门店逐店核查动因"]),
        ("14 · 团队动作与复盘节奏", ["周例会：品类&系列金额复盘", "DSR 拜访成交率周度跟踪", "月复盘：沉淀可复制打法", "季度 PTS 得分拆解与改进"]),
        ("15 · 风险点排查清单", ["低价/窜货/假货风险", "串码/账务冲减导致的负库存", "重复进表/时点口径不一致", "KOC 网点流失预警"]),
        ("16 · 数据治理建议", ["统一经销商编码（主/关联客户簇）", "建立经销商主数据字典", "关键指标口径说明写入报告页尾", "DSR-经销商编码映射校验"]),
        ("17 · 附录：数据来源说明", [
            "新家园进货销售表：可用金额=库存",
            "经销商PTS计分卡：月/季/年得分与排名",
            "经销商分销能力：网点/覆盖率/KOC/终端等级",
            "DSR分销能力：DSR人员维度门店执行数据",
            "城市经理分销能力：区域对标背景数据",
        ]),
        ("18 · 附录：客户簇编码", bucket_codes),
    ]
    for title, items in action_pages:
        pages.append(
            f"""
            <section class="page">
              <h2>{title}</h2>
              <ul class="ul">{''.join([f'<li>{str(x)}</li>' for x in items])}</ul>
            </section>
            """
        )

    body = "\n".join(pages)

    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <title>{dealer_name}（{dealer_code}）经销商白皮书</title>
  <style>
    :root{{
      --navy:#1a3a5c; --navy-d:#122840; --navy-l:#234b73;
      --orange:#f59e0b; --blue:#0ea5e9; --green:#10b981; --red:#ef4444;
      --bg:#eef2f7; --paper:#ffffff; --text:#1e293b; --muted:#64748b;
      --line:#e2e8f0;
    }}
    *{{box-sizing:border-box}}
    body{{margin:0;background:var(--bg);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Microsoft YaHei','PingFang SC',sans-serif;color:var(--text);font-size:14px;line-height:1.6}}
    .wrap{{max-width:960px;margin:0 auto;padding:28px 16px}}

    /* ── Page card ── */
    .page{{
      background:var(--paper);
      border-radius:2px;
      padding:36px 40px;
      margin:12px 0;
      box-shadow:0 1px 4px rgba(0,0,0,.07),0 6px 20px rgba(0,0,0,.05);
      border-left:4px solid var(--navy);
      position:relative;
    }}
    .page h2{{
      margin:0 0 20px;
      font-size:15px;
      font-weight:700;
      color:var(--navy);
      letter-spacing:.2px;
      padding-bottom:14px;
      border-bottom:1px solid var(--line);
      display:flex;
      align-items:center;
      gap:10px;
    }}
    .page h2::before{{
      content:'';
      display:inline-block;
      width:3px;
      height:16px;
      background:var(--orange);
      border-radius:2px;
      flex-shrink:0;
    }}

    /* ── Grid ── */
    .grid3{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px;margin-bottom:16px}}
    .grid2{{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px;margin-bottom:16px}}

    /* ── KPI Card ── */
    .card{{
      background:#f8fafc;
      border:1px solid var(--line);
      border-radius:4px;
      padding:18px 16px 14px;
      border-top:3px solid var(--navy);
    }}
    .k{{font-size:11px;font-weight:600;color:var(--muted);text-transform:uppercase;letter-spacing:.7px;margin-bottom:8px}}
    .v{{font-size:26px;font-weight:800;color:var(--navy);letter-spacing:-.5px;line-height:1.1}}
    .s{{font-size:11px;color:var(--muted);margin-top:8px}}

    /* ── Utility ── */
    .muted{{color:var(--muted)}}
    .small{{font-size:12px;line-height:1.75}}

    /* ── Embedded chart ── */
    .img{{width:100%;border:1px solid var(--line);border-radius:4px;display:block;margin-bottom:16px}}

    /* ── Stat pills ── */
    .pill{{
      display:inline-flex;align-items:center;gap:6px;
      border:1px solid var(--line);
      padding:5px 14px;
      border-radius:999px;
      font-size:12px;font-weight:600;
      margin:8px 6px 0 0;
      background:#fff;
      color:var(--navy);
    }}
    .kpi-row{{margin-top:10px}}

    /* ── Callout / note ── */
    .note{{
      margin-top:16px;
      border-left:4px solid var(--orange);
      background:#fffbeb;
      padding:12px 16px;
      border-radius:0 4px 4px 0;
      color:#78350f;
      line-height:1.8;
      font-size:13px;
    }}

    /* ── Table ── */
    .tbl{{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:12px}}
    .tbl thead tr{{background:var(--navy);color:#fff}}
    .tbl th{{padding:10px 12px;text-align:left;font-weight:600;font-size:11.5px;letter-spacing:.3px}}
    .tbl td{{padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top}}
    .tbl tbody tr:nth-child(even){{background:#f8fafc}}
    .tbl tbody tr:hover{{background:#eff6ff}}
    .tbl tr:last-child td{{border-bottom:none}}

    /* ── Lists ── */
    .ul{{margin:8px 0 0 20px;line-height:1.9}}
    .ol{{margin:8px 0 0 20px;line-height:1.9}}
    .ul li,.ol li{{margin-bottom:2px}}

    /* ── Cover page ── */
    .cover{{
      background:linear-gradient(140deg,var(--navy-d) 0%,var(--navy) 50%,var(--navy-l) 100%);
      color:#fff;
      border-left:none;
      padding:52px 48px 44px;
      min-height:320px;
      overflow:hidden;
    }}
    .cover::before{{
      content:'';position:absolute;top:-100px;right:-80px;
      width:320px;height:320px;
      background:radial-gradient(circle,rgba(245,158,11,.22) 0%,transparent 65%);
      border-radius:50%;
    }}
    .cover::after{{
      content:'';position:absolute;bottom:-80px;left:35%;
      width:220px;height:220px;
      background:radial-gradient(circle,rgba(14,165,233,.18) 0%,transparent 65%);
      border-radius:50%;
    }}
    .cover-top{{
      font-size:11px;font-weight:700;letter-spacing:2.5px;text-transform:uppercase;
      color:var(--orange);margin-bottom:20px;
      position:relative;
    }}
    .cover-title{{
      font-size:38px;font-weight:800;line-height:1.15;
      margin:0 0 14px;letter-spacing:-.5px;
      position:relative;
    }}
    .cover-sub{{color:rgba(255,255,255,.75);font-size:14px;margin-bottom:6px;position:relative}}
    .cover-meta{{color:rgba(255,255,255,.45);font-size:12px;position:relative}}
    .cover-badges{{margin-top:30px;display:flex;gap:12px;flex-wrap:wrap;position:relative}}
    .badge{{
      background:rgba(255,255,255,.1);
      border:1px solid rgba(255,255,255,.22);
      color:#fff;
      padding:8px 18px;
      border-radius:2px;
      font-weight:700;font-size:13px;
      letter-spacing:.2px;
    }}

    /* ── Responsive ── */
    @media (max-width:860px){{
      .grid3{{grid-template-columns:1fr}}
      .grid2{{grid-template-columns:1fr}}
      .cover{{padding:36px 28px 32px}}
      .page{{padding:24px 20px}}
      .cover-title{{font-size:28px}}
    }}

    /* ── Print ── */
    @media print{{
      body{{background:#fff}}
      .wrap{{max-width:none;padding:0}}
      .page{{box-shadow:none;border-radius:0;margin:0;page-break-after:always;border-left:4px solid var(--navy)}}
      .page:last-child{{page-break-after:auto}}
      .cover{{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
    }}
  </style>
</head>
<body>
  <div class="wrap">
    {body}
  </div>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=20, help="最多生成多少个经销商（默认 20；0=全量）")
    ap.add_argument("--codes", type=str, default="", help="仅生成指定经销商编码（逗号分隔）")
    args = ap.parse_args()

    print("Loading inventory detail…")
    inv_detail = _load_inventory_detail()

    # 全体经销商聚合，用于分位对标
    inv_all_agg = inv_detail.groupby("经销商编码", dropna=False, as_index=False)[
        ["库存额", "进货额", "销售额出厂价"]
    ].sum()

    dealer_master = _load_dealer_master()
    pts_df = _load_pts()
    dsr_df = _load_dsr()
    dealer_dist_df = _load_dealer_dist()
    city_mgr_df = _load_city_mgr()
    print(f"  pts={len(pts_df)} rows, dsr={len(dsr_df)} rows, dealer_dist={len(dealer_dist_df)} rows, city_mgr={len(city_mgr_df)} rows")

    # 待生成经销商列表：以进货表出现的编码为主
    codes_all = (
        [c for c in inv_all_agg["经销商编码"].astype(str).str.strip().tolist() if c]
        if "经销商编码" in inv_all_agg.columns
        else []
    )

    want_codes = []
    if args.codes.strip():
        want_codes = [x.strip() for x in args.codes.split(",") if x.strip()]
    else:
        want_codes = codes_all

    if args.limit and args.limit > 0:
        want_codes = want_codes[: args.limit]

    print(f"Generating HTML: {len(want_codes)} dealers → {OUT_DIR}")

    for i, code in enumerate(want_codes, 1):
        name = ""
        if not dealer_master.empty and "客户编码" in dealer_master.columns and "客户名称" in dealer_master.columns:
            hit = dealer_master[dealer_master["客户编码"].astype(str).str.strip().str.upper() == str(code).strip().upper()]
            if not hit.empty:
                name = str(hit.iloc[0]["客户名称"]).strip()
        if not name:
            # fallback: use inventory name
            sub = inv_detail[inv_detail["经销商编码"].astype(str).str.strip().str.upper() == str(code).strip().upper()]
            if not sub.empty:
                name = str(sub.iloc[0]["客户名称"]).strip()
        if not name:
            name = code

        bucket = _dealer_bucket_codes(dealer_master, str(code))
        html = build_dealer_report_html(
            dealer_code=str(code).strip(),
            dealer_name=name,
            bucket_codes=bucket,
            inv_detail=inv_detail,
            inv_all_agg=inv_all_agg,
            pts_df=pts_df,
            dsr_df=dsr_df,
            dealer_dist_df=dealer_dist_df,
            city_mgr_df=city_mgr_df,
            dealer_master=dealer_master,
        )
        fn = f"{_safe_filename(name)}_{str(code).strip()}_经销商白皮书.html"
        out = OUT_DIR / fn
        out.write_text(html, encoding="utf-8")
        if i % 10 == 0 or i == len(want_codes):
            print(f"  {i}/{len(want_codes)} done")

    print("Done.")


if __name__ == "__main__":
    main()

