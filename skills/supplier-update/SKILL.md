---
name: supplier-update
description: 供应商资质自动更新工作流。当用户提到供应商更新、供应商资质检查、到期供应商扫描、住建部资质查询、供应商库维护、安全许可证到期检查、或执行"供应商工作流"时使用本技能。
---

# 供应商资质自动更新工作流 v5.4

## 概述

自动化扫描腾讯云文档供应商 Excel，筛选即将到期的供应商，通过**住建部官网**实时查询最新安全生产许可证有效期，并更新到 Excel 文件中。同时生成**双渠道日志**（本地 + 腾讯云文档）。

```
腾讯云文档 → 筛选30天内到期+已过期 → 住建部官网查询 → 更新Excel → 三渠道通知
```

**通知渠道**：本地日志 + 腾讯云文档持续更新 + 邮件通知

---

## 快速开始

### 安装依赖

```bash
pip install pandas openpyxl patchright
```

### 执行更新（测试3家）

```bash
python /root/.openclaw/workspace/skills/supplier-update/scripts/supplier_update_workflow.py --limit 3
```

### 执行完整更新

```bash
python /root/.openclaw/workspace/skills/supplier-update/scripts/supplier_update_workflow.py
```

---

## 工作流程

| 步骤 | 说明 |
|------|------|
| 1. 导出Excel | 从腾讯云文档导出供应商库（本地已有则跳过） |
| 2. 筛选供应商 | 筛选30天内到期 + 已过期的供应商（有数据才查） |
| 3. 住建部查询 | Playwright 自动化浏览器查询住建部官网 |
| 4. 更新Excel | 将新有效期写入 `*_updated.xlsx` |
| 5. 持续日志 | 腾讯云文档追加本次批次记录 |
| 6. 邮件通知 | 发送HTML格式更新报告至 52644253@qq.com |

---

## 核心脚本

### supplier_update_workflow.py（主流程 v5.2）

```bash
# 测试模式（只查3家）
python supplier_update_workflow.py --limit 3

# 完整模式（查所有）
python supplier_update_workflow.py

# 指定天数
python supplier_update_workflow.py --days 60
```

### mohurd_scraper.py（住建部爬虫）

```python
from mohurd_scraper import query_mohurd_single, query_mohurd_batch

# 查询单家公司
date = query_mohurd_single("江苏八达路桥有限公司")
print(date)  # -> "2026-06-23" 或 None

# 批量查询
results = query_mohurd_batch(["公司A", "公司B", "公司C"])
```

---

## 邮件配置

| 配置项 | 值 |
|--------|-----|
| 发件邮箱 | hxyseujs@163.com |
| SMTP服务器 | smtp.163.com |
| 端口 | 465 (SSL) |
| 收件人 | 52644253@qq.com |

配置文件：`/root/.openclaw/workspace/skills/supplier-update/assets/email_config.json`

---

## 输出渠道

### 1. 本地日志

| 文件 | 路径 |
|------|------|
| JSON日志 | `/root/.openclaw/workspace-supplier/logs/update_log_*.json` |

### 2. 腾讯云文档（持续更新）

- **复用同一文档**，每次运行追加新批次
- **文档URL**：`https://docs.qq.com/sheet/XgHCUPXHAiLi`
- **配置文件**：`/root/.openclaw/workspace-supplier/log_doc_config.json`（记录当前文档ID和写入位置）
- **文档结构**：

| 列 | 内容 |
|---|---|
| 序号 | 全局编号 |
| 执行时间 | 本批次执行时间 |
| 公司名称 | 供应商全称 |
| 原日期 | 更新前日期 |
| 新日期 | 更新后日期 |
| 状态 | 已更新 / 待处理 |
| 批次 | 批次号（如20260402_165632） |

---

## 腾讯云文档配置

- **源文件ID**：`XvdgsQTsPGdZ`
- **文件名称**：合格供应商库_202603061624
- **本地缓存**：`/root/.openclaw/workspace/supplier_full.xlsx`

---

## Excel 列索引

- **列3 (D)**：供应商名称
- **列20 (U)**：安全许可证有效期

---

## 工作流逻辑

```
┌─────────────────────────────────────────────┐
│ 读取Excel                                      │
└─────────────────┬─────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────┐
│ 筛选：30天内到期 + 已过期（有数据才查）        │
└─────────────────┬─────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────┐
│ 住建部官网查询                               │
└─────────────────┬─────────────────────────────┘
                  │
        ┌─────────┴─────────┐
        ▼                   ▼
┌───────────────┐   ┌───────────────┐
│ 有结果        │   │ 无结果        │
│ → 更新Excel   │   │ → 记录日志    │
│ → 记录日志    │   │   等人工处理   │
└───────────────┘   └───────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────┐
│ 双渠道日志                                   │
│   本地: JSON + TXT                          │
│   腾讯: 新建Sheet文档                        │
└─────────────────────────────────────────────┘
```

---

## 住建部爬虫详解

### 技术方案

使用 **patchright**（Playwright Python 版本）实现浏览器自动化：

1. **启动无头浏览器**：Chromium headless 模式
2. **访问查询页面**：https://zlaq.mohurd.gov.cn/fwmh/bjxcjgl/fwmh/pages/construction_safety/qyaqscxkz/qyaqscxkz
3. **展开搜索面板**：点击 `a.btn.show`
4. **填写查询条件**：在 `#qymc` 输入框填写公司名称
5. **触发搜索**：点击 `a.btn.search`
6. **提取结果**：从 `common-table` 结构中获取有效期结束日期

### 页面结构

```
DIV.content container
  DIV.tablebox table-dotted
    DIV.table-contbox
      TABLE.common-table
        TBODY
          TR.common-table-tr
            TD.common-table-td > SPAN (日期)
```

### 注意事项

- 页面使用**字体加密**，公司名称、证号等显示为乱码
- **有效期结束日期为明文**，可直接提取
- 每次查询约需 15-20 秒
- 建议设置 2-3 秒间隔，避免对网站造成压力

---

## 依赖

```bash
pip install pandas openpyxl patchright
```

patchright 会自动安装 Chromium 浏览器。

---

## 目录结构

```
supplier-update/
├── SKILL.md                          # 本文档
├── assets/
│   └── email_config.json            # 邮件配置模板
├── references/
│   ├── workflow.md                   # 工作流详解
│   └── email_config.md              # 邮件配置说明
└── scripts/
    ├── supplier_update_workflow.py   # 主流程脚本 (v5.2)
    ├── mohurd_scraper.py             # 住建部爬虫
    ├── email_sender.py              # 邮件发送模块
    └── supplier_auto_update_v5.py    # 旧版脚本 (保留)
```

---

## 测试日志

### 2026-04-02 测试结果 (v5.2)

```
[1/3] 江苏丰润建设有限公司: 2026-05-01 -> 2029-04-30 ✅
[2/3] 江苏天同建设发展有限公司: 2026-04-09 -> 2026-04-08 ✅
[3/3] 无锡市市政设施建设工程有限公司: 2026-04-06 -> 2026-04-05 ✅

腾讯云文档日志: https://docs.qq.com/sheet/DWGdPREZjSEhadGFS
```

### 历史测试 (v5.1)

```
[1/5] 江苏丰润建设有限公司: 2026-05-01 -> 2029-04-30 ✅
[2/5] 江苏天同建设发展有限公司: 2026-04-09 -> 2026-04-08 ✅
[3/5] 无锡市市政设施建设工程有限公司: 2026-04-06 -> 2026-04-05 ✅
[4/5] 江苏通达建设集团有限公司: 2026-04-30 -> 2029-04-30 ✅
[5/5] 江苏烨瑞建设工程有限公司: 2026-04-07 -> 2029-04-06 ✅
```

---

## 故障排除

### 1. 住建部网站无法访问

**症状**：查询返回 None 或超时

**可能原因**：
- 网站维护或宕机
- 网络连接问题
- IP 被限制

**解决方案**：
- 检查网络连接
- 稍后重试

### 2. Excel 文件未找到

**症状**：`FileNotFoundError: 未找到 Excel 文件`

**解决方案**：
- 确认本地缓存路径 `/root/.openclaw/workspace/supplier_full.xlsx` 存在
- 或确保腾讯云文档可访问

### 3. 查询返回 None

**症状**：公司名称正确但返回 None

**可能原因**：
- 该公司确实没有安全许可证数据
- 公司名称与官网记录不完全一致

**解决方案**：
- 记录到日志，人工处理
- 尝试使用公司全称查询

### 4. 腾讯云文档写入失败

**症状**：本地日志成功但腾讯云文档报错

**可能原因**：
- mcporter 命令执行失败
- sheet_id 获取失败

**解决方案**：
- 检查 mcporter 配置
- 查看错误日志

---

## 更新日志

### v5.4 (2026-04-02)

- ✅ **邮件通知**：集成邮件发送功能
- ✅ 三渠道通知：本地日志 + 腾讯云文档 + 邮件
- ✅ 邮件发送至 52644253@qq.com

### v5.3 (2026-04-02)

- ✅ **持续日志**：复用同一腾讯云文档，每次追加新批次
- ✅ 配置文件 `log_doc_config.json` 记录文档状态
- ✅ 支持断点续传（验证文档有效性后复用）

### v5.2 (2026-04-02)

- ✅ 增加腾讯云文档日志同步功能
- ✅ 双渠道日志：本地 + 腾讯云文档
- ✅ 每次运行自动创建新的腾讯文档

### v5.1 (2026-04-02)

- ✅ 修复住建部爬虫（正确提取日期）
- ✅ 整合腾讯云文档导出功能
- ✅ 实现完整工作流
- ✅ 生成更新日志

### v5.0 (2026-04-02)

- ✅ 整合真实住建部爬虫（Playwright）
- ✅ 替换模拟数据逻辑
- 📝 更新文档

### v4.0 (2026-03-24)

- 流式分块读取优化
- 列投影减少内存占用
- 批量查询合并请求
