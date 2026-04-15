#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
招标邮件通知发送模块
复用 supplier-update 的 email_sender.py SMTP 配置
"""

import smtplib
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
from datetime import datetime

# ── 配置路径（复用 supplier-update 的邮件配置）─────────────────────────────

def get_email_config_path():
    """获取邮件配置文件路径"""
    candidates = [
        Path("/root/.openclaw/workspace/skills/supplier-update/assets/email_config.json"),
        Path("/root/.openclaw/workspace/skills/changzhou-bid-search/email_config.json"),
    ]
    for p in candidates:
        if p.exists():
            return p
    return candidates[0]


def load_email_config():
    """加载邮件配置"""
    path = get_email_config_path()
    if not path.exists():
        raise FileNotFoundError(f"邮件配置文件不存在: {path}，请先配置 {path}")
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def send_email(subject, body, recipients=None, html=False):
    """
    发送邮件

    Args:
        subject: 邮件主题
        body: 邮件正文
        recipients: 收件人列表（默认使用配置文件中的通知列表）
        html: 是否使用 HTML 格式

    Returns:
        bool: 发送成功返回 True
    """
    config = load_email_config()
    smtp = config['smtp']
    notif = config.get('notification', {})

    if recipients is None:
        recipients = notif.get('default_recipients', [])
        if isinstance(recipients, str):
            recipients = [recipients]

    if not recipients:
        print("  [WARN] 无收件人，跳过邮件发送")
        return False

    msg = MIMEMultipart()
    msg['From'] = smtp['email']
    msg['To'] = ', '.join(recipients)
    prefix = notif.get('subject_prefix', '【常州招标】')
    msg['Subject'] = f"{prefix} {subject}"

    content_type = 'html' if html else 'plain'
    msg.attach(MIMEText(body, content_type, 'utf-8'))

    try:
        if smtp.get('use_ssl', True):
            server = smtplib.SMTP_SSL(smtp['server'], smtp['port'])
        else:
            server = smtplib.SMTP(smtp['server'], smtp['port'])
        server.login(smtp['email'], smtp['app_password'])
        server.sendmail(smtp['email'], recipients, msg.as_string())
        server.quit()
        print(f"  邮件已发送至 {len(recipients)} 个收件人: {', '.join(recipients)}")
        return True
    except Exception as e:
        print(f"  邮件发送失败: {e}")
        return False


def build_bid_email_body(items):
    """
    构建招标通知 HTML 邮件正文。

    Args:
        items: 符合资质要求的项目列表，每个元素包含:
            title, area, date, budget, duration, bid_method, qual_check,
            deposit, site, content, contact, section_id 等字段

    Returns:
        str: HTML 正文
    """
    today = datetime.now().strftime('%Y-%m-%d')
    n = len(items)

    # 按日期倒序排列
    items_sorted = sorted(items, key=lambda x: x.get('date', ''), reverse=True)

    rows_html = ""
    for i, it in enumerate(items_sorted, 1):
        budget = it.get('budget', '—')
        deposit = it.get('deposit', '—')
        duration = it.get('duration', '—')
        bid_method = it.get('bid_method', '—')
        qual_check = it.get('qual_check', '—')
        contact = it.get('contact', it.get('agent', '—'))
        site = it.get('site', '—')
        content = (it.get('content', '') or '—')[:80]

        bg = '#f9f9f9' if i % 2 == 0 else '#ffffff'
        rows_html += f"""
        <tr style="background-color: {bg};">
          <td style="padding: 10px; border: 1px solid #ddd; text-align: center; font-weight: bold;">{i}</td>
          <td style="padding: 10px; border: 1px solid #ddd;">
            <strong>{it.get('project_name', it.get('title', '—'))}</strong><br/>
            <span style="color: #666; font-size: 12px;">{it.get('area', '')} · {it.get('date', '')}</span>
          </td>
          <td style="padding: 10px; border: 1px solid #ddd;">{budget}</td>
          <td style="padding: 10px; border: 1px solid #ddd;">{duration}</td>
          <td style="padding: 10px; border: 1px solid #ddd;">{deposit}</td>
          <td style="padding: 10px; border: 1px solid #ddd;">{bid_method}</td>
          <td style="padding: 10px; border: 1px solid #ddd;">{qual_check}</td>
          <td style="padding: 10px; border: 1px solid #ddd; font-size: 13px;">{contact}</td>
        </tr>"""

    html = f"""
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333; max-width: 900px; margin: 0 auto;">
    <div style="background: linear-gradient(135deg, #028090, #02C39A); padding: 20px 30px; border-radius: 8px 8px 0 0;">
        <h2 style="color: white; margin: 0;">🏗️ 常州市政招标公告通知</h2>
        <p style="color: rgba(255,255,255,0.85); margin: 5px 0 0;">
            推送时间: {datetime.now().strftime('%Y-%m-%d %H:%M')} &nbsp;|&nbsp; 本期共 {n} 条符合条件的招标公告
        </p>
    </div>

    <div style="padding: 20px 30px; background: #f5f5f5; border-radius: 0 0 8px 8px;">
        <p style="color: #666; font-size: 14px;">
            以下招标公告已根据贵单位资质进行筛选匹配。如需查看完整公告或下载招标文件，请联系招标代理或访问
            <a href="http://ggzy.xzsp.changzhou.gov.cn" target="_blank">常州市公共资源交易中心</a>。
        </p>
    </div>

    <table style="width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px;">
        <thead>
            <tr style="background-color: #028090; color: white;">
                <th style="padding: 12px 8px; border: 1px solid #ddd; text-align: center; width: 30px;">#</th>
                <th style="padding: 12px 8px; border: 1px solid #ddd; text-align: left;">项目名称 / 地区</th>
                <th style="padding: 12px 8px; border: 1px solid #ddd; text-align: center;">控制价</th>
                <th style="padding: 12px 8px; border: 1px solid #ddd; text-align: center;">工期</th>
                <th style="padding: 12px 8px; border: 1px solid #ddd; text-align: center;">保证金</th>
                <th style="padding: 12px 8px; border: 1px solid #ddd; text-align: center;">评标办法</th>
                <th style="padding: 12px 8px; border: 1px solid #ddd; text-align: center;">资格审查</th>
                <th style="padding: 12px 8px; border: 1px solid #ddd; text-align: left;">联系方式</th>
            </tr>
        </thead>
        <tbody>
            {rows_html}
        </tbody>
    </table>

    <div style="margin-top: 20px; padding: 15px; background: #fff3cd; border-radius: 5px; border-left: 4px solid #ffc107;">
        <h4 style="color: #856404; margin: 0 0 5px;">⚠️ 注意事项</h4>
        <ul style="color: #856404; margin: 0; padding-left: 20px;">
            <li>请在投标截止时间前完成投标，投标保证金需提前缴纳。</li>
            <li>资格审查方式为<b>资格后审</b>，开标后由评标委员会审查。</li>
            <li>控制价为招标最高限价，超过控制价的投标将被否决。</li>
            <li>本通知根据贵单位资质条件自动筛选，仅供参考，招标人不对筛选结果负责。</li>
        </ul>
    </div>

    <hr style="border: none; border-top: 1px solid #ddd; margin: 25px 0 15px;"/>
    <p style="color: #999; font-size: 12px; text-align: center;">
        此邮件由常州招标监控系统自动发送 · 常州市公共资源交易中心<br/>
        发送时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} &nbsp;|&nbsp;
        <a href="http://ggzy.xzsp.changzhou.gov.cn" style="color: #028090;">访问官网</a> &nbsp;|&nbsp;
        <a href="https://docs.qq.com/sheet/DWHZyYmd1eE1kWGp0" style="color: #028090;">腾讯文档</a>
    </p>
</body>
</html>"""
    return html


def send_bid_notifications(matched_items, recipients=None, dry_run=False):
    """
    发送招标公告邮件通知。

    Args:
        matched_items: 符合资质要求的项目列表（来自 qualification_matcher 匹配结果）
        recipients: 收件人邮箱列表（默认使用配置）
        dry_run: True 则仅打印，不实际发送

    Returns:
        bool: 发送成功返回 True
    """
    if not matched_items:
        print("  [INFO] 无匹配项目，跳过邮件发送")
        return False

    today = datetime.now().strftime('%Y-%m-%d')
    subject = f"【新增招标】{len(matched_items)} 条公告符合贵单位资质条件 ({today})"

    if dry_run:
        print(f"  [DRY RUN] 邮件主题: {subject}")
        print(f"  [DRY RUN] 收件人: {recipients}")
        for it in matched_items:
            print(f"  [DRY RUN]   - {it.get('date')} | {it.get('project_name', it.get('title',''))[:40]}")
        return True

    html_body = build_bid_email_body(matched_items)
    return send_email(subject, html_body, recipients=recipients, html=True)


# ── 测试 ──────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    test_items = [
        {
            'project_name': '朱林邻里中心装修改造项目施工总承包工程',
            'area': '金坛区', 'date': '2026-04-03',
            'budget': '437.27 万元', 'duration': '180日历天',
            'bid_method': '合理低价法', 'qual_check': '资格后审',
            'deposit': '8 万元', 'contact': '招标人 0519-85661852',
            'site': '江苏省常州市金坛区朱林镇永兴路 55 号',
        },
        {
            'project_name': '官圩港南站引水河、管网及安装工程',
            'area': '溧阳市', 'date': '2026-04-02',
            'budget': '250 万元', 'duration': '45日历天',
            'bid_method': '合理低价法', 'qual_check': '资格后审',
            'deposit': '5 万元', 'contact': '招标人 15366820221',
            'site': '溧阳市高新区',
        },
    ]
    print("测试邮件内容构建...")
    html = build_bid_email_body(test_items)
    print(f"HTML 长度: {len(html)} bytes")
    print("发送测试邮件（dry_run=True）...")
    send_bid_notifications(test_items, recipients=['test@example.com'], dry_run=True)
