# 邮件配置说明

## 配置文件格式

```json
{
  "smtp": {
    "server": "smtp.163.com",
    "port": 465,
    "email": "your_email@163.com",
    "app_password": "your_16_char授权码",
    "use_ssl": true
  },
  "notification": {
    "default_recipients": [
      "recipient1@example.com",
      "recipient2@example.com"
    ],
    "subject_prefix": "[供应商更新]"
  }
}
```

## 字段说明

### smtp 配置

| 字段 | 类型 | 说明 |
|------|------|------|
| `server` | string | SMTP 服务器地址 |
| `port` | int | SMTP 端口号（465 或 587） |
| `email` | string | 发件人邮箱地址 |
| `app_password` | string | 邮箱授权码（不是登录密码） |
| `use_ssl` | bool | 是否使用 SSL 加密 |

### notification 配置

| 字段 | 类型 | 说明 |
|------|------|------|
| `default_recipients` | array | 默认收件人列表 |
| `subject_prefix` | string | 邮件主题前缀 |

## 常用邮箱 SMTP 配置

### 163 邮箱

```json
{
  "server": "smtp.163.com",
  "port": 465,
  "use_ssl": true
}
```

**授权码获取**：163 邮箱 → 设置 → POP3/SMTP/IMAP → 开启服务 → 获取授权码

### QQ 邮箱

```json
{
  "server": "smtp.qq.com",
  "port": 465,
  "use_ssl": true
}
```

**授权码获取**：QQ 邮箱 → 设置 → 账户 → POP3/IMAP/SMTP/Exchange/CardDAV/CalDAV服务 → 开启 → 获取授权码

### Gmail

```json
{
  "server": "smtp.gmail.com",
  "port": 587,
  "use_ssl": false
}
```

**注意**：Gmail 需要开启"低安全性应用访问权限"或使用应用专用密码

## 常见问题

### 1. 授权码错误

**错误**：`535 authentication failed`

**解决**：确认使用的是**授权码**而非登录密码

### 2. 连接超时

**错误**：`Connection timed out`

**解决**：
- 检查网络连接
- 确认 SMTP 端口未被防火墙阻止
- 尝试更换端口（587）

### 3. SSL 错误

**错误**：`SSL_CTX_check_flags`

**解决**：某些服务器不支持 SSLv3，确认 `use_ssl: true` 且端口为 465

## 测试邮件

执行以下代码测试配置：

```python
from email_sender import send_supplier_update_report

test_updates = [{
    'company': '测试公司',
    'cert_type': 'safety',
    'old_date': '2025-01-01',
    'new_date': '2027-01-01',
    'source': '住建部官网'
}]

success = send_supplier_update_report(test_updates, "测试报告")
print("发送成功" if success else "发送失败")
```
