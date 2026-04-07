## 经销商白皮书（HTML）生成器

该目录用于存放 **经销商白皮书**（经销商经营状况报告）的 HTML 输出。

### 生成方式

在仓库根目录执行：

```bash
python "jxsbps/generate_dealer_html.py" --limit 50
```

默认会读取：
- `xfx_inventory_tool/新家园进货销售刘杨20260323114325.xlsx`（库存/进销/可用金额）
- `xfx_inventory_tool/经销商.xlsx`（客户编码/主客户编码/名称等主数据）
- `baipishu/data/经销商PTS计分卡*.xlsx`（如存在，用于得分/排名）
- `baipishu/data/DSR分销能力*.xlsx`（如存在，用于覆盖/KOC/拜访等）

输出：
- `jxsbps/output_html/*.html`（每个经销商一份单文件 HTML，可直接浏览器打开/打印）

