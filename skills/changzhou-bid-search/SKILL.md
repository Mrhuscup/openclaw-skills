---
AIGC:
    ContentProducer: Minimax Agent AI
    ContentPropagator: Minimax Agent AI
    Label: AIGC
    ProduceID: "00000000000000000000000000000000"
    PropagateID: "00000000000000000000000000000000"
    ReservedCode1: 3044022048761be8b7c1d27abe9601f84585b535ffb6234b5601c94c75b9ed24b780a33b022074db30de833c07140e41d7b182749f7bac560e900893b492c944b9d008eea371
    ReservedCode2: 3045022065341af797d4cf556b162e1fae36ec0da9d3134786519263b661cb34706918fd022100f92119cb4a76bf43f636ee629f6c0fb7bbb31c08e763a7ae91738be421f9177c
description: 常州市公共资源交易中心（ggzy.xzsp.changzhou.gov.cn）招标信息搜索与抓取技能，适用于江苏八达路桥有限公司（公路工程二级、市政一级）。当用户要搜索常州招标信息、查询公共资源交易公告、抓取评标结果、下载招标文件 PDF 时激活。
name: changzhou-bid-search
---

# 常州市公共资源交易中心招标搜索技能

> **更新时间：2026-04-08**  
> 网站于 2026-04 全面改版。本文档描述改版后的接口和抓取方案。
>
> **2026-04-08 修复**：标段编号正则支持 `Bxxx（标段编号）` 格式；同步改用 CSV 写入（`set_range_value_by_csv`）+ 验证重试机制；扫描范围扩大至 2000 行。

---

## 一、网站结构

### 列表页（JS 动态渲染）
- **URL**：`/jyzx/001001/tradeInfonew.html?category={code}`
- **说明**：`category` 参数仅用于页面显示分类筛选，服务器实际返回全部混合数据，需在本地按 `cat` 后3位过滤

### 详情页（静态 HTML，可直接抓取）
- **URL**：`/jyzx/001001/{主类}/{子类}/{YYYYMMDD}/{uuid}.html`
- **获取方式**：通过 AJAX POST 到 `/czggzyweb/frontPageRedirctAction.action?cmd=pageRedirect` 获取实际路径

### PDF 附件（无需登录，直接下载）
- **URL**：`/czggzyweb/WebbuilderMIS/attach/downloadZtbAttach.jspx?attachGuid={guid}&appUrlFlag=ztb007&siteGuid={siteGuid}`
- **说明**：附件链接嵌入在详情页 HTML 的 `onclick="ztbfjyz(...)"` 中
- **过滤规则**：只下载含"招标文件正文"的 PDF，跳过"招标公告"等次要附件

### PDF 字段提取（pymupdf）
- **工具**：Python `pymupdf` 库（需提前安装）
- **安装**：`pip install pymupdf --break-system-packages`
- **说明**：从"招标文件正文.pdf"中提取投标关键字段，基于 PDF 章节编号（2.1/3.1 等）精确定位

---

## 二、分类代码（category_code）

### cat12 结构解析
```
001001 | 001 | 001
  ↑      ↑     ↑
 主类别  子类别  类型

主类别（6位）:
  001001 = 建设工程
  001002 = 交通工程
  001003 = 水利工程
  001004 = 政府采购
  001005 = 土地矿产
  001006 = 国有产权
  001009 = 其他交易

子类（9位）:
  001 = 招标公告/资审公告
  005 = 中标候选人/评标结果公示
  006 = 中标结果公告
  008 = 中标合同
  009 = 合同履行及变更

类型末3位（与9位主类对应）:
  001 = 施工
  002 = 服务
  004 = 货物
```

### 常用 category_code

| code | 名称 | 备注 |
|:---|:---|:---|
| `001001001001` | **建设工程-招标公告-施工** | 默认 |
| `001001001002` | 建设工程-招标公告-服务 | |
| `001001001004` | 建设工程-招标公告-货物 | |
| `001001001` | 建设工程-招标公告/资审公告 | 含施工+服务+货物，需本地过滤 |
| `001001005` | 建设工程-中标候选人公示 | |
| `001001006` | 建设工程-中标结果公告 | |
| `001001008` | 建设工程-中标合同 | |
| `001001009` | 建设工程-合同履行及变更 | |

---

## 三、列表页字段

从列表页 `<tr>` 行提取：

| 字段 | HTML 位置 | 说明 |
|:---|:---|:---|
| `cat` | onclick 第1参数 | 12位或9位，如 `001001001001` |
| `uuid` | onclick 第2参数 | 36位 UUID |
| `title` | `<a title="...">` | 项目/公告标题 |
| `area` | 第3列 | 地区，如"溧阳市" |
| `date` | 第4列 | 发布日期 YYYY-MM-DD |

---

## 四、腾讯文档同步（`--sync`）

### 功能说明
每次运行自动计算 7 天滑动窗口，增量同步到腾讯在线表格。

```
用法: python3 scrape_v2.py --sync [category]
```

| 参数 | 说明 |
|:---|:---|
| `--sync` | 开启腾讯文档同步 |
| `--dry-run` | 模拟运行，不实际写入 |

**同步策略**：
- 读取腾讯文档当前全部数据（section_id 唯一标识）
- 按日期分三类：保留（≤7天）｜新增（本次有但文档无）｜过期（文档有但本次无）
- 清空数据区（`clear_range_cells`，start_row=1 保留表头）+ CSV 写入（`set_range_value_by_csv`，比 JSON cells 格式更稳定）
- 写入后自动验证 section_id 计数，不一致时自动重试
- 过期项目的本地 PDF 文件夹同步删除

**API 调用次数**：1次 GET（2000行扫描）+ 1次 clear + 1次 CSV写入 + 1次验证 = 4次/每次运行

---

## 五、邮件通知（`--notify`）

### 功能说明
根据施工单位资质数据库（`companies/companies.json`），将招标公告按资质要求匹配后，通过邮件精准推送。

```
用法: python3 scrape_v2.py --sync --notify [category]
```

**资质匹配逻辑**：
- 解析项目 `qual_req` 字段，提取资质类别（如"建筑工程施工总承包"）和最低等级（如"三级"）
- 支持"以上"（≥满足）和"及以上"（＝或更高）
- 支持多种文本格式：标准格式和"施工总承包不分行业XXX三级"类格式
- 公司满足 AND 条件（同时具备多类别要求）才推送

**资质等级顺序**：特级 > 一级 > 二级 > 三级

**公司数据库**：`companies/companies.json`
```json
{
  "companies": [
    {
      "name": "公司名称",
      "email": "bid@example.com",
      "qualifications": [
        {"category": "建筑工程施工总承包", "level": "二级"},
        {"category": "市政公用工程施工总承包", "level": "三级"}
      ]
    }
  ]
}
```

**邮件内容**：HTML 表格，包含项目名称、地区、发布日期、控制价、工期、保证金、评标办法、资格审查、联系方式。

**邮件配置**：复用 `supplier-update/assets/email_config.json` 的 SMTP 配置。

---

## 六、详情页抓取流程

### Step 1：获取详情页 URL（AJAX）
```python
POST /czggzyweb/frontPageRedirctAction.action?cmd=pageRedirect
参数:
  infoid     = uuid
  siteGuid   = 7eb5f7f1-9041-43ad-8e13-8fcb82ea831a
  categorynum = cat

响应: {"custom": "/jyzx/001001/001001001/001001001001/20260401/{uuid}.html", ...}
```

### Step 2：抓取详情页 HTML
直接 GET 返回的 URL 即可，包含完整字段和 PDF 附件链接。

### Step 3：解析字段
详情页分两类结构：
1. **表格结构**（中标结果公告等）：通过 `<td align=...>` 键值对提取
2. **正文结构**（招标公告正文）：通过关键词定位提取

---

## 七、执行脚本

### 持久积累模式（默认）
脚本每次运行自动维护一份**持续积累 Excel**（`outputs/常州招标_持续积累.xlsx`），逻辑如下：

1. **读取**本地已有记录（按 section_id 去重）
2. **合并**本次抓取的新条目（标题/日期/字段增量更新）
3. **清理**发布时间距今 >7 天的过期条目 + 对应本地 PDF 文件夹
4. **写入**同一份持久 Excel（固定路径，不按日期命名）
5. **同步**腾讯文档时基于这份完整记录全量重写

```
outputs/
├── 常州招标_持续积累.xlsx      ← 持续积累（每次运行增量合并）
├── 20260408_xxxxxxxx/         ← PDF 文件（每个项目一个文件夹）
└── 20260408_yyyyyyyy/
    └── 招标文件正文.pdf
```

### 基本用法

### 基本用法
```bash
cd /workspace/skills/changzhou-bid-search/scripts
python3 scrape_v2.py [--sync] [--notify] [--dry-run] [category]
```

| 参数 | 默认值 | 说明 |
|:---|:---|:---|
| `category` | `001001001001` | 建设工程-招标公告-施工 |
| `--sync` | 关闭 | 开启腾讯文档同步 |
| `--notify` | 关闭 | 开启邮件资质匹配推送 |
| `--dry-run` | 关闭 | 模拟运行，不实际写入/发送 |

> 注意：`--sync` 和 `--notify` 可单独或同时使用，均支持 `--dry-run` 预览。

### 示例
```bash
# 默认：抓取招标公告-施工，7天窗口，写入持续积累Excel
python3 scrape_v2.py

# 同步腾讯文档（基于持续积累Excel全量同步）
python3 scrape_v2.py --sync

# 同步腾讯文档 + 发送邮件通知
python3 scrape_v2.py --sync --notify

# 模拟运行（预览合并结果，不实际写入）
python3 scrape_v2.py --sync --dry-run

# 中标候选人公示
python3 scrape_v2.py --sync 001001005

# 中标结果公告
python3 scrape_v2.py --sync 001001006
```

### 执行流程（6步）
1. **Step 1** 抓取常州列表页（自动计算近7天窗口）
2. **Step 2** 抓取详情页 + 下载 PDF + 提取字段
3. **Step 3½** 读本地持久Excel → 合并本次结果 → 删除过期条目及PDF → 写入同一持久文件
4. **Step 4** 腾讯文档同步（`--sync`，基于合并后全量数据）
5. **Step 5** 邮件资质匹配推送（`--notify`，`--dry-run` 跳过发送）

---

## 八、输出说明

### 目录结构
```
outputs/
├── 常州招标_持续积累.xlsx      ← 持续积累Excel（固定路径）
├── 20260408_xxxxxxxx/
│   └── 招标文件正文.pdf      # 招标文件（每个项目1个）
└── 20260408_yyyyyyyy/
    └── 招标文件正文.pdf
```

### Excel 字段（24列，来源：招标文件 PDF）

| # | 字段 | 来源 |
|:---:|:---|:---|
| 1 | 序号 | |
| 2 | 项目名称 | PDF 2.1 标段名称 |
| 3 | 地区 | 列表页 |
| 4 | 发布日期 | 列表页 |
| 5 | 分类 | cat 末3位 |
| 6 | 招标人 | PDF 2.1 |
| 7 | 招标代理 | PDF 2.1 |
| 8 | 建设地点 | PDF 2.2 |
| 9 | 建设内容 | PDF 2.3 |
| 10 | 控制价(万元) | PDF 2.6 工程合同估算价 |
| 11 | 评标办法 | PDF 勾选框 + 第三章评标入围章节。综合评估法额外提取 2.3.4 评分组成。格式：`合理低价法，低价排序法（30家）` / `综合评估法，全部入围法，施工组织设计12分，答辩2分，报价81分，信用分5分` |
| 12 | 评标基准价计算方法 | PDF 2.3.2 → 提取评标基准价计算方法（ABC合成法等）、K取值范围、下浮率Δ分类取值范围。格式：`ABC合成法，K（95%-98%），下浮率Δ（6%-12%）` |
| 13 | 资格审查 | PDF 勾选框（☑资格后审）|
| 14 | 投标保证金 | PDF 3.4.1 |
| 15 | 工期 | PDF 要求工期 |
| 16 | 计划开工 | PDF 计划开工日期 |
| 17 | 计划竣工 | PDF 计划竣工日期 |
| 18 | 资质要求 | PDF 3.1 投标人资质 |
| 19 | 项目负责人资质 | PDF 3.2 |
| 20 | 合同价格形式 | PDF 勾选框（☑单价合同）|
| 21 | 投标有效期 | PDF 3.3 |
| 22 | 履约担保 | PDF「合同总价/合同金额…的X%」模式，支持全角/半角％ |
| 23 | 联系方式 | PDF「10.联系方式」章节优先；fallback 为全文首条招标人电话 |
| 24 | 标段编号 | PDF 标段编号 |
| 25 | 来源 | 固定"常州公共资源交易中心" |

> 注：招标文件 PDF 已下载到本地（每个项目一个"招标文件正文.pdf"），PDF列表和PDF数量字段已移除，原始 PDF 文件保留在输出目录中。

---

## 九、自动迭代机制

每次 `--sync` 运行结束后自动进行**字段质量复盘**：

1. 对本次抓取的每条记录，从 PDF 原文核查关键字段（budget、评标办法、资格审查、工期、资质要求等 12 项）
2. 对比提取值与 PDF 原文，标记 ✅/⚠️/❌/N/A
3. 统计本次质量评分（正确提取数 / 应提取总数）
4. 将复盘结果追加写入 `scripts/changelog.md`，持续积累

复盘结果在每次运行末尾输出：
```
[字段质量复盘] ✅ 无问题 | 21/22
```

**查看 changelog**：`cat scripts/changelog.md`

---

## 十、已知限制

| 状态 | 说明 |
|:---|:---|
| ⚠️ 列表翻页 | 当前仅抓第1页（15条），需实现分页逻辑 |
| ⚠️ 字段提取不完整 | 部分公告正文字段提取规则需持续优化 |
| ⚠️ 分类过滤 | 列表页 category 参数无效，本地按 cat 后3位过滤 |

---

## 十一、技术参数

| 参数 | 值 |
|:---|:---|
| siteGuid | `7eb5f7f1-9041-43ad-8e13-8fcb82ea831a` |
| AJAX 接口 | `/czggzyweb/frontPageRedirctAction.action?cmd=pageRedirect` |
| PDF 下载 | `/czggzyweb/WebbuilderMIS/attach/downloadZtbAttach.jspx` |
| PDF 字段提取 | Python `pymupdf`（`pip install pymupdf --break-system-packages`）|
| PDF 过滤规则 | 只下载含"招标文件正文"的 PDF，跳过"招标公告"等次要附件 |
| Excel 列数 | 24列（已移除招标控制价、PDF数量、PDF列表三列）|

---

*最后更新：2026-04-03*
