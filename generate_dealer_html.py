#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
经销商经营状况白皮书（HTML）生成器

目标：参考行业调研报告的表达方式，但用公司现有数据输出 15–20 个“分页段落”的经销商单页 HTML。

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
    # 尽量补一个“经销商编码”列（如果原始表有）
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
    p = _pick_latest_xlsx("经销商PTS计分卡", ROOT / "baipishu" / "data")
    if not p:
        return pd.DataFrame()
    return pd.read_excel(p, sheet_name=0, engine="openpyxl").fillna("")


def _load_dsr() -> pd.DataFrame:
    p = _pick_latest_xlsx("DSR分销能力", ROOT / "baipishu" / "data")
    if not p:
        return pd.DataFrame()
    return pd.read_excel(p, sheet_name=0, engine="openpyxl").fillna("")


def _dealer_bucket_codes(dealer_df: pd.DataFrame, dealer_code: str) -> list[str]:
    """同一主客户簇下的所有客户编码（用于把主/关联户的数据合并到“该经销商报告”）。"""
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
    dealer_master: pd.DataFrame,
) -> str:
    # 过滤明细（按经销商编码“簇”）
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

    # PTS（可选）
    pts_info = {}
    if not pts_df.empty:
        # 尝试按 经销商编码/名称对齐
        cand_cols = [c for c in ["经销商编码", "客户编码", "客户编号", "经销商编号", "经销商名称", "客户名称"] if c in pts_df.columns]
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
            for k in ["客户等级", "季度得分", "季度排名", "省", "市", "客户经理编码", "客户经理工号", "客户经理名称"]:
                if k in r:
                    pts_info[k] = r.get(k)

    # DSR（可选，按经销商编码聚合）
    dsr_info = {}
    if not dsr_df.empty:
        code_cols = [c for c in ["经销商编码", "客户编码", "经销商编号", "客户编号"] if c in dsr_df.columns]
        hit = pd.DataFrame()
        for c in code_cols:
            s = dsr_df[c].astype(str).str.strip().str.upper()
            if dealer_code and (s == dealer_code.upper()).any():
                hit = dsr_df[s == dealer_code.upper()]
                break
        if not hit.empty:
            # 聚合
            cols_try = {
                "整体网点数": ["整体网点数", "网点总数", "门店数"],
                "活跃网点": ["活跃网点", "活跃网点数"],
                "KOC网点数": ["KOC网点数", "KOC门店数"],
                "KOC销售额": ["KOC销售额", "KOC销量", "KOC销售"],
                "拜访门店数": ["拜访门店数", "拜访门店", "拜访数"],
                "拜访成交率": ["拜访成交率", "成交率"],
                "覆盖率": ["覆盖率", "覆盖率(%)"],
            }

            def pick_first(keys):
                for k in keys:
                    if k in hit.columns:
                        return k
                return None

            agg = {}
            for out_k, cands in cols_try.items():
                col = pick_first(cands)
                if not col:
                    continue
                v = pd.to_numeric(hit[col], errors="coerce").fillna(0).sum()
                agg[out_k] = float(v)
            dsr_info = agg

    # 画图（Matplotlib）
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    plt.rcParams.update({"font.sans-serif": ["SimHei", "Microsoft YaHei", "STHeiti"], "axes.unicode_minus": False})

    # 图1：品类库存 TOP
    fig1 = plt.figure(figsize=(7.2, 3.2))
    ax = fig1.add_subplot(111)
    if not cat.empty:
        ax.barh(cat["品类"].astype(str).iloc[::-1], cat["库存额"].astype(float).iloc[::-1], color="#ef4444")
        ax.set_title("品类库存 TOP10（库存额=可用金额）", fontsize=11)
    else:
        ax.text(0.5, 0.5, "无可用品类数据", ha="center", va="center")
        ax.axis("off")
    img_cat = _b64_png(fig1)
    plt.close(fig1)

    # 图2：进销存概览
    fig2 = plt.figure(figsize=(7.2, 2.8))
    ax2 = fig2.add_subplot(111)
    ax2.bar(["进货额", "销售额(出厂)", "库存额"], [in_sum, sale_sum, inv_sum], color=["#3b82f6", "#10b981", "#ef4444"])
    ax2.set_title("进销存概览（当期汇总）", fontsize=11)
    img_kpi = _b64_png(fig2)
    plt.close(fig2)

    # 报告日期
    report_date = datetime.now().strftime("%Y-%m-%d")

    # 15-20 个分页段落（打印时每段一页）
    pages = []

    pages.append(
        f"""
        <section class="page cover">
          <div class="cover-top">经销商经营状况白皮书 &nbsp;·&nbsp; 2025</div>
          <div class="cover-title">{dealer_name or "—"}</div>
          <div class="cover-sub">经销商编码：{dealer_code} &nbsp;|&nbsp; 客户簇编码数：{len(bucket_codes)}</div>
          <div class="cover-meta">生成日期：{report_date} &nbsp;·&nbsp; 数据口径：新家园进货销售表（可用金额=库存）</div>
          <div class="cover-badges">
            <span class="badge">库存额 &nbsp;{_fmt_money(inv_sum)}</span>
            <span class="badge">销售额 &nbsp;{_fmt_money(sale_sum)}</span>
            <span class="badge">负库存行 &nbsp;{neg_rows}</span>
          </div>
        </section>
        """
    )

    pages.append(
        f"""
        <section class="page">
          <h2>01 · 结论摘要（Executive Summary）</h2>
          <div class="grid3">
            <div class="card"><div class="k">库存合计</div><div class="v">{_fmt_money(inv_sum)}</div><div class="s">分位：{_pct(pct_inv)}</div></div>
            <div class="card"><div class="k">销售额（出厂）</div><div class="v">{_fmt_money(sale_sum)}</div><div class="s">分位：{_pct(pct_sale)}</div></div>
            <div class="card"><div class="k">库销比</div><div class="v">{(f'{turnover_ratio:.1f}x' if turnover_ratio is not None else '—')}</div><div class="s">销售=0 时不计算</div></div>
          </div>
          <div class="note">
            <b>解读建议：</b>库存高且库销比高 → 积压风险；负库存行多 → 口径/时点/重复进表需核；品类集中度高 → 重点品类控进货与去化动作更有效。
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
          <div class="small muted">提示：优先关注“库存额高且销售低”的系列；若出现大量负库存，请先核对规则起算日/重复进表/冲减口径。</div>
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
                <li>按库存额排序锁定 TOP3 品类与 TOP10 系列；对“库存高/销售低”做去化动作。</li>
                <li>对库销比高的品类设置进货闸门，优先消化库存再补货。</li>
                <li>若负库存行多：先核“起算日/重复进表/冲减”，再谈补货或去化。</li>
              </ol>
            </div>
            <div class="card">
              <div class="k">动作建议（本月）</div>
              <ol class="ol">
                <li>建立“品类-系列”周度复盘：金额、动销、异常行、责任人、截止日期。</li>
                <li>对重点品类做陈列/促销资源申请并跟踪验收。</li>
                <li>把库存与动销指标纳入经销商月度例会，形成闭环。</li>
              </ol>
            </div>
          </div>
        </section>
        """
    )

    # 能力（DSR）页
    pages.append(
        f"""
        <section class="page">
          <h2>06 · 分销覆盖与门店活跃（若有 DSR 能力数据）</h2>
          <div class="grid3">
            <div class="card"><div class="k">整体网点数</div><div class="v">{_fmt_money(dsr_info.get('整体网点数') if dsr_info else None)}</div></div>
            <div class="card"><div class="k">活跃网点</div><div class="v">{_fmt_money(dsr_info.get('活跃网点') if dsr_info else None)}</div></div>
            <div class="card"><div class="k">KOC 网点数</div><div class="v">{_fmt_money(dsr_info.get('KOC网点数') if dsr_info else None)}</div></div>
          </div>
          <div class="grid3">
            <div class="card"><div class="k">KOC 销售额</div><div class="v">{_fmt_money(dsr_info.get('KOC销售额') if dsr_info else None)}</div></div>
            <div class="card"><div class="k">拜访门店数</div><div class="v">{_fmt_money(dsr_info.get('拜访门店数') if dsr_info else None)}</div></div>
            <div class="card"><div class="k">拜访成交率</div><div class="v">{_fmt_money(dsr_info.get('拜访成交率') if dsr_info else None)}</div></div>
          </div>
          <div class="small muted">说明：若该页均为 “—”，表示当前仓库的 DSR 能力表无法与该经销商编码对齐或数据缺失。</div>
        </section>
        """
    )

    # PTS 页
    pages.append(
        f"""
        <section class="page">
          <h2>07 · 经营评分与等级（若有 PTS 计分卡）</h2>
          <div class="grid3">
            <div class="card"><div class="k">客户等级</div><div class="v">{pts_info.get('客户等级','—')}</div></div>
            <div class="card"><div class="k">季度得分</div><div class="v">{_fmt_money(pts_info.get('季度得分'))}</div></div>
            <div class="card"><div class="k">季度排名</div><div class="v">{pts_info.get('季度排名','—')}</div></div>
          </div>
          <div class="grid2">
            <div class="card"><div class="k">归属客户经理</div><div class="v">{pts_info.get('客户经理名称') or pts_info.get('客户经理编码') or '—'}</div></div>
            <div class="card"><div class="k">地区</div><div class="v">{(pts_info.get('省') or '—')} {(pts_info.get('市') or '')}</div></div>
          </div>
          <div class="small muted">说明：PTS 的字段取决于表结构，若该页为空需要提供“经销商编码/名称”可对齐字段。</div>
        </section>
        """
    )

    # 行业对标页（文本化占位，便于后续加财务/渠道数据）
    pages.append(
        """
        <section class="page">
          <h2>08 · 行业对标（参考调研报告框架）</h2>
          <div class="card">
            <div class="k">当前可对标的维度</div>
            <ul class="ul">
              <li><b>库存结构</b>：品类集中度、TOP 品类占比、负库存异常。</li>
              <li><b>动销与库销</b>：库存/销售关系、停滞风险识别。</li>
              <li><b>覆盖与活跃</b>：网点覆盖、KOC 渗透、拜访与成交（需 DSR 数据对齐）。</li>
            </ul>
          </div>
          <div class="card">
            <div class="k">仍缺失的关键指标（需要新增数据源）</div>
            <ul class="ul">
              <li>毛利/净利、费用结构、人效、现金流、应收账款与逾期。</li>
              <li>渠道结构（传统/现代/量贩/即时零售等）与渠道贡献。</li>
              <li>库存周转天数、动销天数（需期初/期末或出入库口径）。</li>
            </ul>
          </div>
        </section>
        """
    )

    # 9-16：行动手册（拆成多页以形成 15–20 页长度）
    action_pages = [
        ("09 · 本周必做清单", ["锁定 TOP3 品类与 TOP10 系列", "对停滞/高库销比系列制定去化动作", "负库存先核口径再下结论"]),
        ("10 · 进货闸门与补货策略", ["积压品类：先去化再补货", "畅销品类：补货小步快跑", "新品：试销-复盘-再铺货"]),
        ("11 · 终端动销抓手（可验收）", ["陈列位提升（拍照验收）", "促销资源申请（ROI 复盘）", "KOC 门店渗透（名单+节奏）"]),
        ("12 · 团队动作与复盘节奏", ["周例会：品类&系列金额复盘", "责任人：到门店/到货架/到照片", "月复盘：沉淀可复制打法"]),
        ("13 · 风险点排查清单", ["低价/窜货/假货风险", "串码/账务冲减导致的负库存", "重复进表/时点口径不一致"]),
        ("14 · 数据治理建议", ["统一经销商编码（主/关联客户簇）", "建立经销商主数据字典", "关键指标口径说明写入报告页尾"]),
        ("15 · 附录：数据来源说明", ["新家园进货销售表：可用金额=库存", "PTS 计分卡：得分与排名", "DSR 能力表：覆盖/KOC/拜访"]),
        ("16 · 附录：客户簇编码", bucket_codes),
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

