# A 股产业链卡点周报自动化

这个模块把“产业链卡点企业周报”部署到 GitHub Actions。它不依赖本地 Codex 常驻运行，每周日 08:30（Asia/Shanghai）自动生成报告并通过 SMTP 发送邮件。

## GitHub Secrets

在仓库 `Settings -> Secrets and variables -> Actions -> Secrets` 配置：

| Secret | 用途 |
|---|---|
| `ASTOCKANA_SMTP_HOST` | SMTP 主机，例如 `smtp.163.com` |
| `ASTOCKANA_SMTP_PORT` | SMTP SSL 端口，例如 `465` |
| `ASTOCKANA_SMTP_USER` | 发件邮箱 |
| `ASTOCKANA_SMTP_AUTH_CODE` | SMTP 授权码或应用专用密码 |
| `ASTOCKANA_REPORT_RECIPIENT` | 收件邮箱，例如 `u8044657@anu.edu.au` |
| `ASTOCKANA_TUSHARE_TOKEN` | Tushare token，用于行情区间 |

不要把 `.env`、SMTP 授权码、Tushare token 或任何 API key 提交到公开仓库。

## GitHub Variables

可选，在仓库 `Settings -> Secrets and variables -> Actions -> Variables` 配置：

| Variable | 默认值 |
|---|---|
| `ASTOCKANA_MAIL_FROM_NAME` | `A-Share Supply Chain Weekly` |
| `ASTOCKANA_MAIL_SUBJECT_PREFIX` | `A-Share Supply Chain Bottleneck Weekly` |
| `ASTOCKANA_DISABLE_EMAIL` | `false` |

如果只想测试报告生成、不发送邮件，可以把 `ASTOCKANA_DISABLE_EMAIL=true`。

## 输出格式约束

候选表固定为 7 列：

| 产业链层级/卡点环节 | 公司/股票代码 | 排序原因 | 一周新增证据 | 行情/PEG | 主要风险 | 待验证事实/研究优先级 |
|---|---|---|---|---|---|---|

第一列直接使用 `产业链层级/卡点环节` 格式，不显示“产业链层级：”或“卡点环节：”前缀。
第二列公司/股票代码在 Excel 和 HTML 邮件中链接到同花顺个股页，URL 格式为 `https://stockpage.10jqka.com.cn/股票六位代码/`。

正文不显示收件人行、内部邮件标题行，也不显示“完整可筛选版本见附件...”之类附件说明句。表格直接出现在“核心候选公司横向对比表”标题后。

行情和 PEG 使用 Tushare，但表格中不显示数据源字样，不显示年份、交易日期、成交额或近 60 交易日区间。E 列只展示最近交易日收盘价、月涨跌幅和 PEG。

## 运行方式

自动运行：GitHub Actions 每周日 00:30 UTC 触发，对应北京时间/上海时间每周日 08:30。

手动触发：进入 `Actions -> A-Share Supply Chain Bottleneck Weekly -> Run workflow`。

## 文件

- `.github/workflows/a-share-supply-chain-bottleneck.yml`
- `scripts/a_share_supply_chain_bottleneck.py`
- `requirements-bottleneck.txt`

后续对本地自动化的表格、邮件、证据来源或发送规则调整，也应同步更新这些 GitHub 文件。
