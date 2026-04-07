#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
将四个新数据表的内容注入到现有的三个经销商白皮书 HTML 中。
替换/新增：页06（分销覆盖）、07（终端等级）、08（PTS评分）、09（DSR明细）、10（区域对标）
"""

from __future__ import annotations
import base64, io, re
from pathlib import Path
import pandas as pd

HERE = Path(__file__).resolve().parent

# ── 加载四个表 ────────────────────────────────────────────────────────────────

def _pick_latest(prefix: str) -> Path | None:
    cands = sorted(HERE.glob(f"{prefix}*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return cands[0] if cands else None

def load_pts() -> pd.DataFrame:
    p = _pick_latest("经销商PTS计分卡")
    return pd.read_excel(p, engine="openpyxl").fillna("") if p else pd.DataFrame()

def load_dsr() -> pd.DataFrame:
    p = _pick_latest("DSR分销能力")
    return pd.read_excel(p, engine="openpyxl").fillna(0) if p else pd.DataFrame()

def load_dealer_dist() -> pd.DataFrame:
    p = _pick_latest("经销商分销能力")
    if not p: return pd.DataFrame()
    df = pd.read_excel(p, engine="openpyxl").fillna(0)
    for c in ["经销商编码", "经销商名称", "大区", "区域", "城市", "客户经理工号", "客户经理名称", "经销商等级"]:
        if c in df.columns: df[c] = df[c].astype(str).str.strip()
    return df

def load_city_mgr() -> pd.DataFrame:
    p = _pick_latest("城市经理分销能力")
    if not p: return pd.DataFrame()
    df = pd.read_excel(p, engine="openpyxl").fillna(0)
    for c in ["大区", "区域", "城市经理工号", "城市经理名称"]:
        if c in df.columns: df[c] = df[c].astype(str).str.strip()
    return df

# ── 工具函数 ──────────────────────────────────────────────────────────────────

def _fmt(x) -> str:
    try:
        if x is None or x == "" or x == 0: return "—"
        v = float(x)
        if v == 0: return "—"
        return f"{v:,.0f}"
    except Exception:
        return str(x) if x else "—"

def _b64_png(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("ascii")

def _match_dealer(df: pd.DataFrame, code: str, name: str, code_col="经销商编码", name_col="经销商名称") -> pd.DataFrame:
    if df.empty: return pd.DataFrame()
    if code_col in df.columns:
        s = df[code_col].astype(str).str.strip().str.upper()
        hit = df[s == code.upper()]
        if not hit.empty: return hit
    if name_col in df.columns and name:
        hit = df[df[name_col].astype(str).str.strip() == name]
        if not hit.empty: return hit
    return pd.DataFrame()

# ── 构建新 pages HTML ─────────────────────────────────────────────────────────

def build_new_pages(code: str, name: str,
                    pts_df: pd.DataFrame, dsr_df: pd.DataFrame,
                    dist_df: pd.DataFrame, cm_df: pd.DataFrame) -> str:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({"font.sans-serif": ["SimHei", "Microsoft YaHei", "STHeiti"], "axes.unicode_minus": False})

    # PTS
    pts_info: dict = {}
    pts_hit = _match_dealer(pts_df, code, name, "经销商编码", "经销商名称")
    if not pts_hit.empty:
        r = pts_hit.iloc[0].to_dict()
        for k in ["客户等级", "当月得分", "季度得分", "季度排名", "年度得分", "年度排名",
                  "大区", "区域", "城市", "客户经理编码", "客户经理姓名"]:
            if k in r and r[k] not in ("", None, 0, "0"): pts_info[k] = r[k]

    # DSR 聚合
    dsr_info: dict = {}
    dsr_rows = _match_dealer(dsr_df, code, name, "经销商编码", "经销商名称")
    if not dsr_rows.empty:
        for col in ["整体网点数", "活跃网点", "KOC网点数", "KOC销售额", "拜访门店数", "拜访成交门店数", "销售额"]:
            if col in dsr_rows.columns:
                dsr_info[col] = float(pd.to_numeric(dsr_rows[col], errors="coerce").fillna(0).sum())
        if "拜访成交率" in dsr_rows.columns:
            v = pd.to_numeric(dsr_rows["拜访成交率"], errors="coerce").mean()
            if pd.notna(v): dsr_info["拜访成交率"] = float(v)

    # 经销商分销能力
    dist_info: dict = {}
    dist_hit = _match_dealer(dist_df, code, name)
    if not dist_hit.empty:
        r = dist_hit.iloc[0]
        for f in ["经销商等级", "整体网点数", "KOC网点数", "KOC销售额", "市场容量", "覆盖率(%)",
                  "活跃网点", "品牌分销指数", "复用指数", "经销商分销能力",
                  "本月金标店数", "本月银标店数", "本月铜标店数", "本月基础店数", "本月不达标数",
                  "KOC-火机品类网点数", "KOC-辣味零食品类网点数", "KOC-爆珠网点数", "KOC-剃须刀品类网点数",
                  "大区", "区域", "城市", "客户经理名称"]:
            if f in r.index and r[f] not in ("", None): dist_info[f] = r[f]

    # 城市经理
    cm_info: dict = {}
    region = str(dist_info.get("区域") or pts_info.get("区域", "")).strip()
    if region and "区域" in cm_df.columns:
        cm_rows = cm_df[cm_df["区域"].astype(str).str.strip() == region]
        if not cm_rows.empty:
            cm_r = cm_rows.iloc[0]
            for f in ["城市经理名称", "整体网点数", "KOC网点数", "KOC销售额", "覆盖率(%)", "分销能力", "市场容量"]:
                if f in cm_r.index: cm_info[f] = cm_r[f]

    # ── 图：终端等级分布 ──
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
        fig, ax = plt.subplots(figsize=(5.5, 3.0))
        colors = ["#f59e0b", "#94a3b8", "#b45309", "#64748b", "#ef4444"]
        bars = ax.bar(list(grade_vals.keys()), list(grade_vals.values()), color=colors)
        ax.set_title("终端门店等级分布", fontsize=11)
        ax.tick_params(labelsize=9)
        for bar, v in zip(bars, grade_vals.values()):
            if v > 0:
                ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.3,
                        f"{int(v)}", ha="center", va="bottom", fontsize=9)
        fig.tight_layout()
        img_grade = _b64_png(fig)
        plt.close(fig)

    # ── 图：品类KOC网点 ──
    img_koc = ""
    koc_map = {k: float(dist_info.get(v, 0) or 0) for k, v in [
        ("火机", "KOC-火机品类网点数"), ("辣味零食", "KOC-辣味零食品类网点数"),
        ("爆珠", "KOC-爆珠网点数"), ("剃须刀", "KOC-剃须刀品类网点数"),
    ]}
    koc_map = {k: v for k, v in koc_map.items() if v > 0}
    if koc_map:
        fig, ax = plt.subplots(figsize=(5.5, 3.0))
        ax.barh(list(koc_map.keys()), list(koc_map.values()), color="#0ea5e9")
        ax.set_title("品类 KOC 网点数", fontsize=11)
        ax.tick_params(labelsize=9)
        fig.tight_layout()
        img_koc = _b64_png(fig)
        plt.close(fig)

    # ── 辅助值 ──
    _total_net = float(dist_info.get("整体网点数") or dsr_info.get("整体网点数", 0) or 0)
    _active_net = float(dist_info.get("活跃网点") or dsr_info.get("活跃网点", 0) or 0)
    _active_rate = f"活跃率 {_active_net/_total_net*100:.0f}%" if _total_net > 0 and _active_net > 0 else ""
    _coverage = dist_info.get("覆盖率(%)", "")
    _coverage_str = f"{float(_coverage):.1f}%" if _coverage not in ("", None, 0) else "—"
    _brand_idx = dist_info.get("品牌分销指数", "—")
    _reuse_idx = dist_info.get("复用指数", "—")
    _dist_cap = dist_info.get("经销商分销能力", "—")
    _mkt_cap = dist_info.get("市场容量", "—")
    _koc_sales = dist_info.get("KOC销售额") or dsr_info.get("KOC销售额")

    _pts_monthly = pts_info.get("当月得分", "—")
    _pts_quarter = pts_info.get("季度得分", "—")
    _pts_q_rank = pts_info.get("季度排名", "—")
    _pts_annual = pts_info.get("年度得分", "—")
    _pts_a_rank = pts_info.get("年度排名", "—")

    # ── DSR明细表 ──
    _dsr_visit_rate = ""
    if dsr_info.get("拜访成交率"):
        v = float(dsr_info["拜访成交率"])
        _dsr_visit_rate = f"{v*100:.1f}%" if v <= 1 else f"{v:.1f}%"
    _dsr_rows_html = ""
    if not dsr_rows.empty:
        show_cols = [c for c in ["DSR名称", "整体网点数", "活跃网点", "KOC网点数", "拜访门店数", "拜访成交门店数", "销售额"] if c in dsr_rows.columns]
        rows_html = []
        for _, row in dsr_rows[show_cols].iterrows():
            cells = "".join(f"<td>{row[c] if row[c] else '—'}</td>" for c in show_cols)
            rows_html.append(f"<tr>{cells}</tr>")
        headers = "".join(f"<th>{c}</th>" for c in show_cols)
        _dsr_rows_html = f"""<table class="tbl"><thead><tr>{headers}</tr></thead><tbody>{"".join(rows_html)}</tbody></table>"""

    # ── 终端等级表 ──
    _grade_html = ""
    if total_grade > 0:
        pcts = {k: f"{v/total_grade*100:.0f}%" for k, v in grade_vals.items()}
        rows = "".join(f"<tr><td>{k}</td><td>{int(v)}</td><td>{pcts[k]}</td></tr>" for k, v in grade_vals.items())
        _grade_html = f"""<table class="tbl" style="max-width:360px">
          <thead><tr><th>等级</th><th>门店数</th><th>占比</th></tr></thead>
          <tbody>{rows}<tr><td><b>合计</b></td><td><b>{int(total_grade)}</b></td><td>100%</td></tr></tbody>
        </table>"""

    # ── 拼接 pages ──
    p06 = f"""
    <section class="page">
      <h2>06 · 分销覆盖与网点能力</h2>
      <div class="grid3">
        <div class="card"><div class="k">整体网点数</div><div class="v">{_fmt(_total_net)}</div><div class="s">{_active_rate}</div></div>
        <div class="card"><div class="k">KOC 网点数</div><div class="v">{_fmt(dist_info.get("KOC网点数") or dsr_info.get("KOC网点数"))}</div><div class="s">KOC 销售额 {_fmt(_koc_sales)}</div></div>
        <div class="card"><div class="k">市场覆盖率</div><div class="v">{_coverage_str}</div><div class="s">市场容量 {_fmt(_mkt_cap)}</div></div>
      </div>
      <div class="grid3">
        <div class="card"><div class="k">品牌分销指数</div><div class="v">{_brand_idx}</div><div class="s">越高品牌铺货越广</div></div>
        <div class="card"><div class="k">复用指数</div><div class="v">{_reuse_idx}</div><div class="s">跨品类复用门店能力</div></div>
        <div class="card"><div class="k">经销商分销能力</div><div class="v">{_dist_cap}</div><div class="s">综合分销能力评分</div></div>
      </div>
      {"<img class='img' src='data:image/png;base64," + img_koc + "'/>" if img_koc else ""}
      <div class="small muted">说明：网点数来自经销商分销能力表；覆盖率 = 整体网点数 / 市场容量 × 100%。</div>
    </section>"""

    p07 = f"""
    <section class="page">
      <h2>07 · 终端门店质量分布</h2>
      {"<img class='img' src='data:image/png;base64," + img_grade + "'/>" if img_grade else "<div class='note'>终端等级数据缺失（经销商分销能力表无对应编码）。</div>"}
      {_grade_html}
      <div class="small muted" style="margin-top:12px">金标→银标→铜标→基础→不达标；提升路径：不达标→基础→铜标优先，KOC 门店着重保金/银。</div>
    </section>"""

    p08 = f"""
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
      <div class="small muted">数据来源：经销商PTS计分卡。季度/年度排名越小越靠前。</div>
    </section>"""

    p09 = f"""
    <section class="page">
      <h2>09 · DSR 分销员执行力明细</h2>
      <div class="grid3">
        <div class="card"><div class="k">拜访门店数</div><div class="v">{_fmt(dsr_info.get("拜访门店数"))}</div></div>
        <div class="card"><div class="k">拜访成交门店</div><div class="v">{_fmt(dsr_info.get("拜访成交门店数"))}</div></div>
        <div class="card"><div class="k">拜访成交率</div><div class="v">{_dsr_visit_rate or "—"}</div></div>
      </div>
      {_dsr_rows_html if _dsr_rows_html else '<div class="note">DSR 明细数据暂无（编码未对齐）。</div>'}
      <div class="small muted" style="margin-top:12px">目标拜访成交率 &gt; 60%；数据来源：DSR分销能力表。</div>
    </section>"""

    _cm_name = cm_info.get("城市经理名称", "—")
    _cm_region = dist_info.get("区域") or pts_info.get("区域", "—")
    _cm_cov = cm_info.get("覆盖率(%)", "")
    _cm_cov_str = f"{float(_cm_cov):.1f}%" if _cm_cov not in ("", None, 0) else "—"

    p10 = f"""
    <section class="page">
      <h2>10 · 区域背景与城市经理对标</h2>
      <div class="grid3">
        <div class="card"><div class="k">所属区域</div><div class="v" style="font-size:18px">{_cm_region}</div><div class="s">城市经理：{_cm_name}</div></div>
        <div class="card"><div class="k">区域整体网点</div><div class="v">{_fmt(cm_info.get("整体网点数"))}</div><div class="s">区域 KOC 网点 {_fmt(cm_info.get("KOC网点数"))}</div></div>
        <div class="card"><div class="k">区域覆盖率</div><div class="v">{_cm_cov_str}</div><div class="s">区域分销能力 {cm_info.get("分销能力","—")}</div></div>
      </div>
      <div class="note">
        <b>对标方向：</b>对比本经销商（覆盖率 {_coverage_str}，分销能力 {_dist_cap}）与区域城市经理汇总数据，识别该经销商在区域中的能力水位。覆盖率/分销能力明显低于区域时，优先做网点开发与 KOC 渗透。
      </div>
    </section>"""

    return p06 + p07 + p08 + p09 + p10


# ── 主逻辑：找 HTML 文件并注入 ────────────────────────────────────────────────

OLD_PAGES_PATTERN = re.compile(
    r'<section class="page">\s*<h2>06[^<]*</h2>.*?(?=<section class="page">\s*<h2>0[789]|<section class="page">\s*<h2>1[0-9]|</div>\s*</body>)',
    re.DOTALL
)

# 更简单：找到页06开始位置，找到最后一个行动页之前，替换掉所有06-10
REPLACE_FROM = re.compile(r'(<section class="page">\s*<h2>06\s·)', re.DOTALL)
REPLACE_TO = re.compile(r'(<section class="page">\s*<h2>(?:09|10|11)\s·)', re.DOTALL)


def patch_html(html_path: Path, new_pages: str) -> bool:
    content = html_path.read_text(encoding="utf-8")
    # Find start of page 06
    m_start = re.search(r'<section class="page">\s*\n\s*<h2>06\s·', content)
    if not m_start:
        print(f"  WARNING: page 06 not found in {html_path.name}, appending before </body>")
        content = content.replace("</div>\n</body>", new_pages + "\n</div>\n</body>")
        html_path.write_text(content, encoding="utf-8")
        return True

    start_pos = m_start.start()
    # Find start of the first action page (09 or higher) after page06
    m_end = re.search(r'<section class="page">\s*\n\s*<h2>(?:09|10|11)\s·', content[start_pos:])
    if m_end:
        end_pos = start_pos + m_end.start()
    else:
        # fallback: replace until </div></body>
        m_body = re.search(r'\n\s*</div>\s*\n</body>', content[start_pos:])
        end_pos = start_pos + m_body.start() if m_body else len(content)

    new_content = content[:start_pos] + new_pages + "\n    " + content[end_pos:]
    html_path.write_text(new_content, encoding="utf-8")
    return True


if __name__ == "__main__":
    print("Loading data tables...")
    pts_df = load_pts()
    dsr_df = load_dsr()
    dist_df = load_dealer_dist()
    cm_df = load_city_mgr()
    print(f"  PTS: {len(pts_df)} rows, DSR: {len(dsr_df)} rows, Dist: {len(dist_df)} rows, CityMgr: {len(cm_df)} rows")

    html_files = list(HERE.glob("*_经销商白皮书.html"))
    if not html_files:
        print("No HTML files found.")
        exit(1)

    for html_path in html_files:
        # Extract dealer code from filename
        m = re.search(r'_(C\d+)_', html_path.name)
        if not m:
            print(f"  SKIP: cannot parse code from {html_path.name}")
            continue
        code = m.group(1)
        # Extract dealer name from filename
        name = html_path.name.split("_")[0]
        print(f"  Patching {html_path.name} (code={code})...")

        new_pages = build_new_pages(code, name, pts_df, dsr_df, dist_df, cm_df)
        patch_html(html_path, new_pages)
        print(f"    Done.")

    print("All HTML files patched.")
