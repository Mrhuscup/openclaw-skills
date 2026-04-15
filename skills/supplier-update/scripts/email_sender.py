#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
邮件发送模块 - 用于供应商更新报告通知
"""

import smtplib
import json
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path
from datetime import datetime


def get_config_path():
    """获取配置文件路径"""
    # 尝试多个可能的路径
    possible_paths = [
        Path("/root/.openclaw/workspace-skills/supplier-update/email_config.json"),
        Path(__file__).parent.parent / "assets" / "email_config.json",
        Path(__file__).parent.parent / "email_config.json",
    ]
    
    for p in possible_paths:
        if p.exists():
            return p
    
    return possible_paths[0]  # 返回第一个作为默认值


def load_config():
    """加载邮件配置"""
    config_path = get_config_path()
    
    if not config_path.exists():
        raise FileNotFoundError(f"配置文件不存在: {config_path}")
    
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def send_email(subject, body, recipients=None, html=False):
    """
    发送邮件
    
    Args:
        subject: 邮件主题
        body: 邮件正文
        recipients: 收件人列表（默认使用配置中的 default_recipients）
        html: 是否使用 HTML 格式
    
    Returns:
        bool: 发送成功返回 True
    """
    config = load_config()
    smtp_config = config['smtp']
    notification_config = config['notification']
    
    if recipients is None:
        recipients = notification_config.get('default_recipients', 
            [notification_config.get('default_recipient')] if notification_config.get('default_recipient') else [])
    
    if isinstance(recipients, str):
        recipients = [recipients]
    
    msg = MIMEMultipart()
    msg['From'] = smtp_config['email']
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = f"{notification_config['subject_prefix']} {subject}"
    
    content_type = 'html' if html else 'plain'
    msg.attach(MIMEText(body, content_type, 'utf-8'))
    
    try:
        if smtp_config.get('use_ssl', True):
            server = smtplib.SMTP_SSL(smtp_config['server'], smtp_config['port'])
        else:
            server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
        
        server.login(smtp_config['email'], smtp_config['app_password'])
        server.sendmail(smtp_config['email'], recipients, msg.as_string())
        server.quit()
        
        print(f"✓ 邮件已发送至 {len(recipients)} 个收件人: {', '.join(recipients)}")
        return True
        
    except Exception as e:
        print(f"✗ 邮件发送失败: {str(e)}")
        return False


def send_supplier_update_report(updates, report_text, recipients=None, report_extra=''):
    """
    发送供应商更新报告
    
    Args:
        updates: 更新记录列表
        report_text: 完整报告文本
        recipients: 收件人列表
        report_extra: 额外报告信息
    
    Returns:
        bool: 发送成功返回 True
    """
    if not updates:
        print("无更新内容，跳过邮件发送")
        return False
    
    today = datetime.now().strftime('%Y-%m-%d')
    subject = f"{today} 更新报告 - {len(updates)} 条记录"
    
    # 生成 HTML 邮件正文
    html_body = f"""
    <html>
    <body style="font-family: Arial, sans-serif; line-height: 1.6;">
        <h2 style="color: #333;">📋 供应商信息更新报告</h2>
        <p style="color: #666;">更新时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}</p>
        
        <div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin: 20px 0;">
            <h3 style="color: #2196F3; margin-top: 0;">📊 更新摘要</h3>
            <p><strong>更新数量:</strong> {len(updates)} 条</p>
            <p><strong>数据来源:</strong> 住建部官网 - 安全生产许可证</p>
            {f'<p><strong>📁 更新版本:</strong> {report_extra}</p>' if report_extra else ''}
        </div>
        
        <h3 style="color: #2196F3;">📝 更新详情</h3>
        <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
            <thead>
                <tr style="background-color: #2196F3; color: white;">
                    <th style="padding: 12px; text-align: left;">公司</th>
                    <th style="padding: 12px;">证书类型</th>
                    <th style="padding: 12px;">原有效期</th>
                    <th style="padding: 12px;">新有效期</th>
                    <th style="padding: 12px;">数据来源</th>
                </tr>
            </thead>
            <tbody>
    """
    
    for update in updates:
        html_body += f"""
                <tr style="background-color: #f9f9f9;">
                    <td style="padding: 10px; border-bottom: 1px solid #ddd;">{update['company']}</td>
                    <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: center;">安全施工许可证</td>
                    <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: center;">{update['old_date']}</td>
                    <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: center; color: #4CAF50; font-weight: bold;">{update['new_date']}</td>
                    <td style="padding: 10px; border-bottom: 1px solid #ddd; text-align: center;">{update['source']}</td>
                </tr>
        """
    
    html_body += """
            </tbody>
        </table>
        
        <div style="background-color: #fff3cd; padding: 15px; border-radius: 5px; margin: 20px 0; border-left: 4px solid #ffc107;">
            <h4 style="color: #856404; margin-top: 0;">⚠️ 提醒</h4>
            <p style="color: #856404; margin: 0;">请核对更新信息，确保数据准确。如有疑问，请访问住建部官网核实。</p>
        </div>
        
        <hr style="border: none; border-top: 1px solid #ddd; margin: 30px 0;">
        <p style="color: #999; font-size: 12px; text-align: center;">
            此邮件由供应商信息自动更新系统发送<br>
            发送时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        </p>
    </body>
    </html>
    """
    
    return send_email(subject, html_body, recipients, html=True)


if __name__ == '__main__':
    # 测试邮件发送
    test_updates = [
        {
            'company': '测试公司',
            'cert_type': 'safety',
            'old_date': '2025-01-01',
            'new_date': '2027-01-01',
            'source': '住建部官网'
        }
    ]
    
    print("发送测试邮件...")
    success = send_supplier_update_report(test_updates, "测试报告")
    
    if success:
        print("✓ 测试邮件发送成功")
    else:
        print("✗ 测试邮件发送失败")
