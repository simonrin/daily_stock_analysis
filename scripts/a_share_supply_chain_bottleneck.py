#!/usr/bin/env python3
"""Generate and email the A-share supply-chain bottleneck weekly report.

This script is designed for GitHub Actions. Secrets are read from environment
variables only; do not commit .env files or real tokens to the repository.
"""

from __future__ import annotations

import html
import json
import mimetypes
import os
import smtplib
import ssl
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from email.header import Header
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


REPORT_DIR = Path("reports")
HEADERS = [
    "产业链层级/卡点环节",
    "公司/股票代码",
    "排序原因",
    "一周新增证据",
    "行情/PEG",
    "主要风险",
    "待验证事实/研究优先级",
]


@dataclass(frozen=True)
class Candidate:
    layer: str
    bottleneck: str
    company: str
    ts_code: str
    reason: str
    weekly_evidence: str
    evidence_strength: str
    risk: str
    validation: str
    priority: str


CANDIDATES: list[Candidate] = [
    Candidate("AI 数据中心电力/液冷", "液冷温控、机房热管理", "英维克", "002837.SZ", "AI 机柜功率密度提升，冷却从成本项变成扩产约束", "AI 数据中心电力、冷却和机柜功率密度继续是扩容前置约束", "强", "液冷价格竞争、AI 收入占比不清", "AI 数据中心订单、液冷收入占比、毛利率", "高"),
    Candidate("AI 数据中心电力/液冷", "精密空调、液冷与数据中心环境控制", "申菱环境", "301018.SZ", "弹性高，若订单兑现更敏感", "水/电/冷却约束继续强化，需跟踪项目制订单兑现", "中强", "估值波动、订单确认滞后", "大客户项目与交付周期", "中高"),
    Candidate("国产 AI 芯片/HBM/封装", "国产 AI 加速器", "寒武纪", "688256.SH", "国产替代政策若推进，算力芯片是最硬约束之一", "国产算力基础设施投资预期仍围绕芯片、封装和生态展开", "中强", "性能生态、客户集中、估值压力", "订单、收入确认、软件生态", "高"),
    Candidate("国产 AI 芯片/HBM/封装", "国产 CPU/GPU/加速计算平台", "海光信息", "688041.SH", "国产算力基础设施核心候选", "国产芯片供给和先进制程约束仍是瓶颈", "中强", "先进制程和供应链限制", "数据中心客户、库存和毛利率", "高"),
    Candidate("先进封装/HBM", "高端封测、Chiplet/先进封装", "长电科技", "600584.SH", "算力芯片扩产不只卡晶圆，也卡封装、测试和良率", "国产 AI 芯片供给约束推高先进封装重要性", "中强", "传统封测拖累、资本开支压力", "先进封装收入占比和客户", "中高"),
    Candidate("高速互连", "高速光模块", "中际旭创", "300308.SZ", "AI 集群带宽约束强，客户认证和良率构成壁垒", "AI 数据中心投资预期继续升温，高速互连仍是核心卡点", "中强", "海外客户集中、估值拥挤", "800G/1.6T 订单和毛利率", "高"),
    Candidate("高速互连", "高速光模块", "新易盛", "300502.SZ", "与海外 AI 客户链条相关度高", "高速光模块代际升级和大客户持续性仍需验证", "中强", "订单波动、技术路线切换", "大客户持续性、产品代际", "高"),
    Candidate("高速 PCB/CCL", "AI 服务器 PCB", "沪电股份", "002463.SZ", "高速互连向 PCB/CCL 扩散，低损耗材料和良率是关键", "AI 网络带宽和集群规模扩张，向低损耗 PCB/CCL 扩散", "中", "估值、客户验证节奏", "AI 服务器板收入占比", "中高"),
    Candidate("半导体设备", "刻蚀、薄膜、清洗等平台型设备", "北方华创", "002371.SZ", "国产晶圆扩产上限取决于关键设备验证和交付", "国产 AI 芯片和先进封装需求凸显设备国产替代必要性", "强", "晶圆厂 capex 波动、核心零部件", "新签订单、验收、毛利率", "高"),
    Candidate("半导体设备", "刻蚀、MOCVD", "中微公司", "688012.SH", "刻蚀工艺验证周期长，替代难度高", "关键设备验证周期长，仍是扩产节奏的硬约束", "强", "客户验收周期、价格竞争", "刻蚀订单和先进节点验证", "高"),
    Candidate("半导体设备", "CVD/ALD 薄膜设备", "拓荆科技", "688072.SH", "先进制程和先进封装都需要薄膜沉积能力", "薄膜沉积设备受先进制程和先进封装双重牵引", "中强", "单一赛道波动、订单集中", "新产品导入和验收", "中高"),
    Candidate("人形机器人执行器", "谐波减速器", "绿的谐波", "688017.SH", "如果机器人进入规模试产，精密传动先于整机卡住", "机器人主题需回到客户定点、良率和量产节奏验证", "中", "量产时点、估值、竞争", "客户定点、量产订单", "中高"),
    Candidate("人形机器人执行器", "丝杠/精密零部件", "贝斯特", "300580.SZ", "丝杠寿命和一致性决定量产可靠性", "样件到量产的可靠性验证是核心", "中", "样件到量产落差", "丝杠订单和良率", "中"),
    Candidate("人形机器人执行器", "电机/执行器", "鸣志电器", "603728.SH", "空心杯/控制电机是关节运动基础", "执行器、电机和控制系统仍需确认客户和量产收入", "中", "海外订单不透明、需求波动", "机器人收入占比", "中"),
    Candidate("稀土永磁", "高性能钕铁硼磁材", "金力永磁", "300748.SZ", "机器人、电动车、风电共同需求，磁材认证周期长", "高端磁材认证周期和稀土供应约束仍需跟踪", "中强", "稀土价格波动、出口政策变化", "订单、价格、海外客户", "中高"),
    Candidate("稀土永磁", "钕铁硼磁材", "中科三环", "000970.SZ", "老牌磁材供应商，受益于高端磁材需求", "高端产品占比决定盈利弹性", "中", "盈利弹性不确定", "高端产品占比", "中"),
    Candidate("电网主设备", "换流阀、继保、直流输电", "许继电气", "000400.SZ", "新能源外送和 AI 负荷都需要电网扩容", "AI 负荷和新能源外送强化电网主设备需求", "强", "招标节奏、回款周期", "国网/南网中标包", "高"),
    Candidate("电网主设备", "GIS/开关设备", "平高电气", "600312.SH", "特高压和主网投资可通过招标验证", "主网投资和特高压招标是硬验证路径", "强", "招标价格下降", "GIS 中标份额和毛利率", "中高"),
    Candidate("电网自动化", "继保、调度、自动化", "国电南瑞", "600406.SH", "确定性强，弹性低于小市值设备商", "新型电力系统建设需要调度、继保和自动化", "强", "估值弹性有限", "新型电力系统订单", "中高"),
    Candidate("低空基础设施", "空管系统、低空监管平台", "莱斯信息", "688631.SH", "低空经济真实卡点在空域管理和监管", "低空经济真实约束在监管平台、通信导航和空域管理", "中强", "地方财政、订单披露不透明", "地方低空平台中标", "中高"),
    Candidate("低空基础设施", "空管/雷达/军工电子", "四川九洲", "000801.SZ", "低空监管和通信监视需要硬件基础设施", "低空监管需要通信、监视和导航基础设施", "中", "主题属性强、业务拆分不清", "低空项目收入占比", "中"),
    Candidate("低空基础设施", "通信导航", "海格通信", "002465.SZ", "低空智联网需要通信导航能力", "通信导航项目中标是主要验证点", "中", "订单确认慢", "通信导航项目中标", "中"),
    Candidate("新能源高端材料", "电解液/新型电解质", "天赐材料", "002709.SZ", "固态/高压体系若推进，材料验证是前置约束", "固态和高压体系仍需公司级客户验证", "中", "传统电解液周期下行", "固态电解质客户验证", "中"),
    Candidate("新能源高端材料", "高端隔膜", "恩捷股份", "002812.SZ", "高安全隔膜仍有技术壁垒，但行业供需承压", "高端膜材料有技术壁垒，但行业价格压力仍大", "中", "价格战、产能过剩", "高端隔膜毛利率和客户", "中"),
]

PUBLIC_SIGNAL_QUERIES = [
    "China AI data center power cooling liquid cooling",
    "China semiconductor equipment advanced packaging HBM",
    "China humanoid robot actuator harmonic reducer",
    "China ultra high voltage grid equipment tender",
    "China low altitude economy air traffic management",
    "China rare earth magnet export robotics",
]


def getenv(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def today_cn() -> date:
    return datetime.now(timezone(timedelta(hours=8))).date()


def trade_dates(days_back: int = 45) -> Iterable[str]:
    current = today_cn()
    for offset in range(days_back + 1):
        yield (current - timedelta(days=offset)).strftime("%Y%m%d")


def fetch_quotes(candidates: list[Candidate]) -> dict[str, str]:
    token = getenv("ASTOCKANA_TUSHARE_TOKEN") or getenv("TUSHARE_TOKEN")
    if not token:
        return {c.ts_code: "行情待复核" for c in candidates}

    try:
        import tushare as ts
    except Exception:
        return {c.ts_code: "行情待复核" for c in candidates}

    pro = ts.pro_api(token)
    wanted = {c.ts_code for c in candidates}
    quotes: dict[str, str] = {}
    quote_dates: dict[str, str] = {}

    for trade_date in trade_dates():
        try:
            frame = pro.daily(trade_date=trade_date)
        except Exception:
            continue
        if frame is None or frame.empty:
            continue
        for _, row in frame.iterrows():
            code = str(row.get("ts_code", ""))
            if code not in wanted or code in quotes:
                continue
            close = float(row.get("close"))
            quotes[code] = f"收盘 {close:.2f}"
            quote_dates[code] = trade_date
        if wanted.issubset(quotes):
            break

    for code in wanted - set(quotes):
        try:
            frame = pro.daily(ts_code=code, start_date="20200101", end_date=today_cn().strftime("%Y%m%d"))
        except Exception:
            continue
        if frame is None or frame.empty:
            continue
        try:
            row = frame.sort_values("trade_date", ascending=False).iloc[0]
            quotes[code] = f"收盘 {float(row.get('close')):.2f}"
            quote_dates[code] = str(row.get("trade_date", ""))
        except Exception:
            continue

    pe_ttm: dict[str, float] = {}
    for code in wanted:
        try:
            frame = pro.daily_basic(ts_code=code, trade_date=quote_dates.get(code, ""), fields="ts_code,trade_date,pe_ttm")
        except Exception:
            continue
        if frame is None or frame.empty:
            continue
        try:
            value = float(frame.iloc[0].get("pe_ttm"))
        except Exception:
            continue
        if value > 0:
            pe_ttm[code] = value

    growth: dict[str, float] = {}
    for code in wanted:
        try:
            frame = pro.fina_indicator(ts_code=code, fields="ts_code,end_date,netprofit_yoy")
        except Exception:
            continue
        if frame is None or frame.empty:
            continue
        frame = frame.sort_values("end_date", ascending=False)
        for _, row in frame.iterrows():
            try:
                value = float(row.get("netprofit_yoy"))
            except Exception:
                continue
            if value > 0:
                growth[code] = value
                break

    out = {}
    for c in candidates:
        q = quotes.get(c.ts_code, "行情待复核")
        pe = pe_ttm.get(c.ts_code)
        yoy = growth.get(c.ts_code)
        peg = f"PEG {pe / yoy:.2f}" if pe and yoy else "PEG 待复核"
        out[c.ts_code] = f"{q}；{peg}"
    return out


def layer_summary() -> list[str]:
    seen: list[str] = []
    for c in CANDIDATES:
        if c.layer not in seen:
            seen.append(c.layer)
    return seen


def rows_for_report(quotes: dict[str, str]) -> list[list[str]]:
    rows = []
    for c in CANDIDATES:
        rows.append([
            f"产业链层级：{c.layer}；卡点环节：{c.bottleneck}",
            f"{c.company}（{c.ts_code}）",
            c.reason,
            f"{c.weekly_evidence}；证据强度：{c.evidence_strength}",
            quotes.get(c.ts_code, "行情待复核"),
            c.risk,
            f"待验证：{c.validation}；优先级：{c.priority}",
        ])
    return rows


def fetch_public_signals(max_per_query: int = 2) -> list[dict[str, str]]:
    """Fetch recent public-source signals without requiring a paid API.

    GDELT is used only as a public article index. The report treats these as
    leads to verify, not as final proof of company fundamentals.
    """
    signals: list[dict[str, str]] = []
    for query in PUBLIC_SIGNAL_QUERIES:
        params = urllib.parse.urlencode({
            "query": query,
            "mode": "artlist",
            "format": "json",
            "maxrecords": str(max_per_query),
            "timespan": "7d",
            "sort": "hybridrel",
        })
        url = f"https://api.gdeltproject.org/api/v2/doc/doc?{params}"
        try:
            request = urllib.request.Request(url, headers={"User-Agent": "AStockAna-BottleneckWeekly/1.0"})
            with urllib.request.urlopen(request, timeout=12) as response:
                data = json.loads(response.read().decode("utf-8", errors="replace"))
        except Exception as exc:
            signals.append({"theme": query, "title": f"公开来源待复核：{exc.__class__.__name__}", "url": "", "date": ""})
            continue
        for article in data.get("articles", [])[:max_per_query]:
            title = str(article.get("title") or "").strip()
            article_url = str(article.get("url") or "").strip()
            seen_date = str(article.get("seendate") or "").strip()
            if title:
                signals.append({"theme": query, "title": title, "url": article_url, "date": seen_date})
    return signals[:12]


def write_xlsx(path: Path, rows: list[list[str]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Candidate Table"
    ws.append(["A-Share Supply Chain Bottleneck Weekly"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    ws.append(HEADERS)
    for row in rows:
        ws.append(row)

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    border = Border(
        left=Side(style="thin", color="9BB7D4"),
        right=Side(style="thin", color="9BB7D4"),
        top=Side(style="thin", color="9BB7D4"),
        bottom=Side(style="thin", color="9BB7D4"),
    )
    for row in ws.iter_rows():
        for cell in row:
            wrap_text = not (cell.column == 3 and cell.row >= 3)
            cell.alignment = Alignment(wrap_text=wrap_text, vertical="top")
            cell.border = border
    for cell in ws[2]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{get_column_letter(len(HEADERS))}{ws.max_row}"
    widths = [30, 20, 30, 34, 28, 24, 32]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    ws2 = wb.create_sheet("Layer Ranking")
    ws2.append(["排序", "产业链卡点层级"])
    for idx, layer in enumerate(layer_summary(), 1):
        ws2.append([idx, layer])
    for row in ws2.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border
    for cell in ws2[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 60
    wb.save(path)


def html_table(rows: list[list[str]]) -> str:
    header = "".join(f"<th>{html.escape(h)}</th>" for h in HEADERS)
    body = []
    for row in rows:
        body.append("<tr>" + "".join(f"<td>{html.escape(str(cell))}</td>" for cell in row) + "</tr>")
    return f"<table><thead><tr>{header}</tr></thead><tbody>{''.join(body)}</tbody></table>"


def render_html(rows: list[list[str]], signals: list[dict[str, str]]) -> str:
    layers = "".join(f"<li>{html.escape(layer)}</li>" for layer in layer_summary())
    table = html_table(rows)
    signal_items = "".join(
        "<li>"
        + html.escape(f"{item.get('date', '')} {item.get('theme', '')}: {item.get('title', '')}".strip())
        + (f" <a href=\"{html.escape(item['url'])}\">source</a>" if item.get("url") else "")
        + "</li>"
        for item in signals
    ) or "<li>公开来源待复核。</li>"
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #1f2937; line-height: 1.55; }}
    h2 {{ color: #174A7C; border-bottom: 1px solid #B8D0EA; padding-bottom: 8px; }}
    table {{ border-collapse: collapse; width: 100%; table-layout: fixed; }}
    th, td {{ border: 1px solid #9BB7D4; padding: 10px; vertical-align: top; word-break: break-word; }}
    th {{ background: #D9EAF7; font-weight: 700; }}
    tr:nth-child(even) td {{ background: #F7FBFF; }}
  </style>
</head>
<body>
  <h2>本周最值得研究的 {len(layer_summary())} 个卡点层级</h2>
  <ol>{layers}</ol>
  <h2>核心候选公司横向对比表</h2>
  {table}
  <h2>过去一周公开来源线索</h2>
  <ol>{signal_items}</ol>
  <h2>需要继续验证的事实清单</h2>
  <ol>
    <li>巨潮资讯和交易所公告：重大合同、中标、产能扩建、调研纪要和风险提示。</li>
    <li>国网/南网、地方低空经济平台、数据中心项目招投标。</li>
    <li>Tushare/交易所行情：收盘价和 PEG 比率；不展示日期、成交额或近 60 交易日区间。</li>
    <li>公司级收入结构：AI、机器人、低空、先进封装等业务真实占比。</li>
  </ol>
  <h2>简明结论</h2>
  <p>本周研究优先级仍应先看真实供给约束层级，再看个股弹性。AI 数据中心电力/液冷、国产算力芯片与先进封装、高速互连、半导体设备、电网主设备的证据强度较高；机器人、低空经济和新能源高端材料需要更多订单和收入结构验证。本文仅为研究优先级，不构成买入、卖出或持有建议。</p>
</body>
</html>
"""


def render_text(rows: list[list[str]], signals: list[dict[str, str]]) -> str:
    lines = [f"本周最值得研究的 {len(layer_summary())} 个卡点层级"]
    lines.extend(f"{idx}. {layer}" for idx, layer in enumerate(layer_summary(), 1))
    lines.append("")
    lines.append("核心候选公司横向对比表")
    lines.append("\t".join(HEADERS))
    lines.extend("\t".join(row) for row in rows)
    lines.append("")
    lines.append("过去一周公开来源线索")
    if signals:
        lines.extend(
            f"{idx}. {item.get('date', '')} {item.get('theme', '')}: {item.get('title', '')} {item.get('url', '')}".strip()
            for idx, item in enumerate(signals, 1)
        )
    else:
        lines.append("公开来源待复核。")
    lines.append("")
    lines.append("本文仅为研究优先级，不构成买入、卖出或持有建议。")
    return "\n".join(lines) + "\n"


def send_email(subject: str, text_body: str, html_body: str, attachments: list[Path]) -> None:
    if getenv("ASTOCKANA_DISABLE_EMAIL", "false").lower() == "true":
        print("EMAIL_DISABLED_BY_ASTOCKANA_DISABLE_EMAIL=true")
        return

    host = getenv("ASTOCKANA_SMTP_HOST")
    port = int(getenv("ASTOCKANA_SMTP_PORT", "465") or "465")
    user = getenv("ASTOCKANA_SMTP_USER")
    password = getenv("ASTOCKANA_SMTP_AUTH_CODE")
    recipient = getenv("ASTOCKANA_REPORT_RECIPIENT")
    from_name = getenv("ASTOCKANA_MAIL_FROM_NAME", "A-Share Supply Chain Weekly")
    missing = [name for name, value in {
        "ASTOCKANA_SMTP_HOST": host,
        "ASTOCKANA_SMTP_USER": user,
        "ASTOCKANA_SMTP_AUTH_CODE": password,
        "ASTOCKANA_REPORT_RECIPIENT": recipient,
    }.items() if not value]
    if missing:
        raise RuntimeError(f"Missing email secrets: {', '.join(missing)}")

    msg = EmailMessage()
    msg["From"] = formataddr((str(Header(from_name, "utf-8")), user))
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(text_body, subtype="plain", charset="utf-8")
    msg.add_alternative(html_body, subtype="html", charset="utf-8")

    for path in attachments:
        ctype, _ = mimetypes.guess_type(path.name)
        if ctype is None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        msg.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)

    with smtplib.SMTP_SSL(host, port, context=ssl.create_default_context(), timeout=30) as smtp:
        smtp.login(user, password)
        smtp.send_message(msg)
    print(f"EMAIL_SENT_TO={recipient}")


def main() -> None:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    report_date = today_cn().isoformat()
    prefix = REPORT_DIR / f"a-share-supply-chain-bottleneck-{report_date}"
    quotes = fetch_quotes(CANDIDATES)
    rows = rows_for_report(quotes)
    signals = fetch_public_signals()

    html_body = render_html(rows, signals)
    text_body = render_text(rows, signals)
    html_path = prefix.with_suffix(".html")
    text_path = prefix.with_suffix(".txt")
    doc_path = Path(str(prefix) + ".doc")
    xlsx_path = prefix.with_suffix(".xlsx")
    html_path.write_text(html_body, encoding="utf-8")
    text_path.write_text(text_body, encoding="utf-8")
    doc_path.write_text(html_body, encoding="utf-8")
    write_xlsx(xlsx_path, rows)

    subject_prefix = getenv("ASTOCKANA_MAIL_SUBJECT_PREFIX", "A-Share Supply Chain Bottleneck Weekly")
    subject = f"{subject_prefix} - {report_date}"
    send_email(subject, text_body, html_body, [xlsx_path, doc_path])


if __name__ == "__main__":
    main()
