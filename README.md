# Payroll2DingTalk

> 工资单推送钉钉OA审批 — 单文件 Streamlit Demo

---

## 功能

用户通过手机号登录 → 上传多个 Excel 工资表 → 预览解析数据 → 一键提交钉钉 OA 审批流程。

## 快速开始

### 1. 克隆仓库

```bash
git clone https://github.com/wanghannew1/Payroll2DingTalk.git
cd Payroll2DingTalk
```

### 2. 创建虚拟环境（推荐 uv）

```bash
# 使用 uv 创建虚拟环境（需先安装 uv）
uv venv

# 激活虚拟环境
source .venv/bin/activate
```

> 如果没有 uv，可以用 `python -m venv .venv` 代替。

### 3. 安装依赖

```bash
# 国内用户可使用清华镜像加速
uv pip install -r requirements.txt

# 或指定镜像源
uv pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 4. 配置环境变量

创建 `.env` 文件（见下方【环境配置】章节）。

### 5. 启动服务

**方式一：手动启动（开发调试）**

```bash
streamlit run demo_app.py
```

浏览器打开 `http://localhost:8501`。

**方式二：Systemd 服务（生产环境，开机自启）**

创建服务文件：

```bash
sudo nano /etc/systemd/system/streamlit-payroll.service
```

写入以下内容（注意替换路径）：

```ini
[Unit]
Description=Streamlit Payroll2DingTalk App
After=network.target

[Service]
Type=simple
User=vod
WorkingDirectory=/home/vod/code/Payroll2DingTalk

# 使用虚拟环境的绝对路径
ExecStart=/home/vod/code/Payroll2DingTalk/.venv/bin/streamlit run /home/vod/code/Payroll2DingTalk/demo_app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true

# 自动重启配置
Restart=always
RestartSec=10

# 日志输出
StandardOutput=append:/home/vod/code/Payroll2DingTalk/streamlit.log
StandardError=append:/home/vod/code/Payroll2DingTalk/streamlit-error.log

[Install]
WantedBy=multi-user.target
```

启用并启动服务：

```bash
sudo systemctl daemon-reload
sudo systemctl enable streamlit-payroll
sudo systemctl start streamlit-payroll

# 查看状态
sudo systemctl status streamlit-payroll
```

**常用命令：**

```bash
# 查看日志
sudo tail -f /home/vod/code/Payroll2DingTalk/streamlit.log

# 重启服务
sudo systemctl restart streamlit-payroll

# 停止服务
sudo systemctl stop streamlit-payroll
```

## 环境配置

从 [钉钉开放平台](https://open-dev.dingtalk.com/) 获取以下配置：

| 配置项 | 获取路径 | 说明 |
|--------|----------|------|
| `DINGTALK_APP_KEY` | 应用开发 → 企业内部应用 → 应用信息 → AppKey | 应用的唯一标识 |
| `DINGTALK_APP_SECRET` | 应用开发 → 企业内部应用 → 应用信息 → AppSecret | 应用的密钥 |
| `DINGTALK_AGENT_ID` | 应用开发 → 企业内部应用 → 应用信息 → AgentId | 钉钉微应用编号 |
| `DINGTALK_PROCESS_CODE` | 钉钉管理后台 → OA审批 → 审批流程 → 流程编码 | OA 审批流程的唯一码 |

创建 `.env` 文件：

```env
DINGTALK_APP_KEY=YOUR_APP_KEY
DINGTALK_APP_SECRET=YOUR_APP_SECRET
DINGTALK_AGENT_ID=YOUR_AGENT_ID
DINGTALK_PROCESS_CODE=YOUR_PROCESS_CODE
```

## 钉钉权限配置

在 [钉钉开放平台](https://open-dev.dingtalk.com/) → 应用开发 → 企业内部应用 → 权限管理 中，必须开通以下权限：

### 必须开通（核心功能依赖）

| 权限点 | 用途 | 对应 API |
|--------|------|----------|
| `qyapi_base` | 获取 Access Token | `oauth2/accessToken`、`gettoken` |
| `qyapi_get_member_by_mobile` | 手机号查 userId | `topapi/v2/user/getbymobile` |
| `qyapi_get_member` | 查用户详情（unionId、deptId） | `topapi/v2/user/get` |
| `Workflow.Instance.Write` | 创建审批实例、授权上传空间 | `processInstances`、`spaces/infos/query` |
| `Storage.UploadInfo.Read` | 获取 OSS 上传凭证 | `uploadInfos/query` |
| `Storage.File.Write` | 提交文件上传确认 | `commit` |

### 建议保留（方便后续扩展）

| 权限点 | 用途 |
|--------|------|
| `Contact.User.Read` | 扩展通讯录功能 |
| `Workflow.Form.Read` | 动态获取审批模板字段 |
| `Workflow.Instance.Read` | 查询已提交审批单状态 |
| `Storage.File.Read` | 读取已上传的文件 |

### 不需要开通（本项目未使用）

| 权限点 | 说明 |
|--------|------|
| `Drive.Space.Read` / `Drive.Space.Write` / `Drive.SpaceManage.Read` | 钉盘 API，本项目使用 Storage API |
| `Storage.DownloadInfo.Read` | 下载文件用，本项目只上传 |
| `snsapi_base` | 网页 OAuth 授权，后端调用不需要 |

> **注意**：上传附件使用的是 **Storage API**（`Storage.UploadInfo.Read` + `Storage.File.Write`），不是 Drive API。如果开通 Drive 权限但文件上传仍然失败，请检查是否调用了正确的 Storage 接口。

## 生产环境补充配置

### 防火墙放行

如果服务器启用了防火墙，需开放 8501 端口：

```bash
sudo ufw allow 8501/tcp
sudo ufw reload
```

### Nginx 反向代理（可选）

如需通过域名访问，可配置 Nginx：

```nginx
server {
    listen 80;
    server_name your-domain.com;

    location / {
        proxy_pass http://127.0.0.1:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

### 日志轮转

避免日志文件无限增长：

```bash
sudo nano /etc/logrotate.d/streamlit-payroll
```

写入：

```
/home/vod/code/Payroll2DingTalk/*.log {
    daily
    rotate 7
    compress
    delaycompress
    missingok
    notifempty
    create 0644 vod vod
}
```

## 配置文件（可选）

项目支持通过 `config.json` 自定义 Excel 解析规则和钉钉审批表格字段，无需修改代码即可适配不同的 Excel 模板或 OA 表单。

### 配置文件用途

- **Excel 列名映射**：当工资表的列名发生变化（如 `"实发合计"` 改为 `"实发工资"`），只需修改配置文件中的 `keywords`
- **钉钉表格字段映射**：当钉钉 OA 审批模板的表格字段名称变化时，修改 `table_field.columns` 中的 `label`
- **单位名称提取**：支持自定义正则表达式匹配单位名称

### 配置文件格式

创建 `config.json`（与 `demo_app.py` 同级）：

```json
{
  "excel": {
    "summary_row_marker": "合计",
    "unit_name_patterns": [
      "(?:单位名称[：:]|[名称][：:]\\s*)(.+?)(?:\\s|$)"
    ],
    "columns": {
      "transfer_total": {
        "keywords": ["转款合计", "转账合计"],
        "label": "转账合计（元）"
      },
      "deduction_total": {
        "keywords": ["扣款合计"],
        "label": "扣款合计（五险一金、单位代理费）"
      },
      "net_total": {
        "keywords": ["实发工资"],
        "label": "实发工资（元）"
      },
      "personal_tax": {
        "keywords": ["个税", "个人所得税"],
        "label": "个人所得税"
      },
      "adjustment": {
        "keywords": ["调差", "差额调整", "调整差额", "工伤差额", "返还差额"],
        "label": "调差"
      },
      "service_fee": {
        "keywords": ["服务费", "代理费", "管理费"],
        "label": "服务费"
      },
      "employer_insurance": {
        "keywords": ["单位缴纳", "单位社保", "单位五险一金"],
        "label": "单位缴纳"
      }
    }
  },
  "table_field": {
    "columns": [
      {"key": "report_name", "label": "报表名称"},
      {"key": "unit_name", "label": "甲方单位项目名称"},
      {"key": "transfer_total", "label": "转账合计（元）"},
      {"key": "deduction_total", "label": "扣款合计（五险一金、单位代理费）"},
      {"key": "net_total", "label": "实发合计（元）"},
      {"key": "tax_and_others", "label": "个人所得税及其他"}
    ]
  },
  "ui": {
    "template_name": "工资发放审批",
    "description": "请上传 Excel 工资表，系统将自动解析数据并提交钉钉 OA 审批流程。"
  }
}
```

### 配置项说明

| 路径 | 说明 |
|------|------|
| `excel.summary_row_marker` | 汇总行的首列标识文本，默认 `"合计"` |
| `excel.unit_name_patterns` | 提取单位名称的正则表达式列表，按顺序匹配 |
| `excel.columns.{key}.keywords` | 查找 Excel 列时匹配的文本列表，按顺序尝试 |
| `excel.columns.{key}.label` | 该列在数据预览中的显示名称 |
| `table_field.columns` | 钉钉 OA 审批表格字段列表，`key` 对应内部数据键，`label` 对应钉钉表单字段名 |

> **重要：`table_field.columns` 中的 `key` 不是随意定义的**，它必须和 `parse_excel()` 函数返回的 dict 中的 key 一一对应。目前可用的 key 如下：
>
> | key | 数据来源 | 说明 |
> |-----|----------|------|
> | `report_name` | Excel 第 1 行 | 报表标题 |
> | `unit_name` | Excel 第 2 行 或 标题行正则提取 | 单位名称 |
> | `transfer_total` | 合计行「转款/转账合计」列 | 转账合计 |
> | `deduction_total` | 合计行「扣款合计」列 | 扣款合计 |
> | `net_total` | 合计行「实发工资」列 | 实发工资 |
> | `personal_tax` | 合计行「个税/个人所得税」列 | 个人所得税 |
> | `adjustment` | 合计行「调差/差额调整」列 | 调差（可正可负） |
> | `service_fee` | 合计行「服务费/代理费」列 | 服务费 |
> | `employer_insurance` | 合计行「单位缴纳/单位社保」列 | 单位缴纳社保公积金 |
> | `tax_and_others` | 程序计算：`transfer_total − deduction_total − net_total` | 个人所得税及其他 |
>
> **缺列处理**：只要列在 `excel.columns` 里定义了，即使某张工资表里没有这个列，程序也不会报错——该列按 **0** 参与计算。这让你可以把所有可能出现的列都预先配置上，适配不同的工资表模板。

### 向后兼容

如果 `config.json` 不存在或解析失败，程序会自动使用内置的默认配置，旧用户无需任何改动即可正常运行。

---

## 配置文件完整指南

以下是一份**可直接使用的完整配置模板**，包含当前项目支持的所有配置项。你可以根据实际工资表结构调整：

```json
{
  "excel": {
    "title_row": 1,
    "unit_name_row": 0,
    "header_start_row": 2,
    "header_row_count": 3,
    "summary_row_marker": "合计",
    "unit_name_patterns": [
      "(?:单位名称[：:]|[名称][：:]\\s*)(.+?)(?:\\s|$)"
    ],
    "title_unit_suffixes": [
      "有限公司", "股份公司", "分公司", "公司", "集团",
      "医院", "卫生院", "诊所",
      "研究院", "研究所", "学院", "大学", "学校",
      "中心", "管委会", "事业部", "处", "局"
    ],
    "title_unit_allowed_chars": "[一-龥A-Za-z0-9（）()·\\-—]",
    "title_unit_patterns": [],
    "columns": {
      "transfer_total": {
        "keywords": ["转款合计", "转账合计"],
        "label": "转账合计（元）"
      },
      "deduction_total": {
        "keywords": ["扣款合计"],
        "label": "扣款合计（五险一金、单位代理费）"
      },
      "net_total": {
        "keywords": ["实发工资"],
        "label": "实发工资（元）"
      },
      "personal_tax": {
        "keywords": ["个税", "个人所得税"],
        "label": "个人所得税"
      },
      "adjustment": {
        "keywords": ["调差", "差额调整", "调整差额", "工伤差额", "返还差额"],
        "label": "调差"
      },
      "service_fee": {
        "keywords": ["服务费", "代理费", "管理费"],
        "label": "服务费"
      },
      "employer_insurance": {
        "keywords": ["单位缴纳", "单位社保", "单位五险一金"],
        "label": "单位缴纳"
      }
    }
  },
  "validation": {
    "enabled": true,
    "strict": true,
    "tolerance": 0.00,
    "write_back_sheet": true,
    "write_back_sheet_name": "验证结果",
    "column_sum_checks": [
      {"column": "deduction_total"},
      {"column": "net_total"},
      {"column": "personal_tax"}
    ],
    "row_formulas": [
      {
        "name": "转款合计 = 扣款合计 + 个人所得税 + 实发工资",
        "lhs": "transfer_total",
        "rhs_plus": ["deduction_total", "personal_tax", "net_total"],
        "rhs_minus": []
      }
    ]
  },
  "table_field": {
    "columns": [
      {"key": "report_name", "label": "报表名称"},
      {"key": "transfer_total", "label": "转账合计（元）"},
      {"key": "deduction_total", "label": "扣款合计（五险一金、单位代理费）"},
      {"key": "net_total", "label": "实发工资（元）"},
      {"key": "tax_and_others", "label": "个人所得税及其他"}
    ]
  },
  "ui": {
    "template_name": "工资发放审批",
    "description": "请上传 Excel 工资表，系统将自动解析数据并提交钉钉 OA 审批流程。"
  }
}
```

### excel 段详解

| 配置项 | 必填 | 说明 |
|--------|------|------|
| `title_row` | 否，默认 1 | 报表标题所在行（1-based）。如工资表第 1 行是标题，填 `1` |
| `unit_name_row` | 否，默认 2 | 单位名称所在行（1-based）。**填 `0` 表示没有独立的单位名称行**，程序会从标题行用正则提取 |
| `header_start_row` | 否，默认 3 | 表头起始行（1-based）。如表头从第 2 行开始，填 `2` |
| `header_row_count` | 否，默认 3 | 表头占几行。多级合并表头（如 3 行）填 `3`，单行表头填 `1` |
| `summary_row_marker` | 否，默认 `"合计"` | 汇总行首列的标识文本。支持中间有空格变体（如 `"合 计"`） |
| `unit_name_patterns` | 否 | 从单位名称行提取单位名的正则列表。按顺序匹配，第一个命中的 group(1) 即为单位名 |
| `title_unit_suffixes` | 否 | 当 `unit_name_row=0` 时，从标题行提取单位名所用的组织后缀列表（如 `"有限公司"`、`"研究院"`） |
| `columns` | **是** | Excel 列定义。每个 key 对应一个列，`keywords` 用于在表头中匹配列位置 |

#### columns 配置规则

`excel.columns` 里定义的每个列，会被程序用来：
1. 在表头区域（`header_start_row` 开始的 `header_row_count` 行）中搜索匹配的列名
2. 找到后在合计行中取出该列的数值
3. 同时该 key 也可以被 `validation` 段和 `table_field` 段引用

**示例**：如果你的工资表里「扣款」列的表头写的是 **"本月扣款合计（五险）"**，而 `keywords` 里配的是 `["扣款合计"]`，程序会命中（因为是子串包含匹配）。但如果表头改成了 **"扣款小计"**，就需要把 `keywords` 改成 `["扣款小计", "扣款合计"]`。

> **重要规则**：`validation` 和 `table_field` 里引用的列 key（如 `personal_tax`），**必须先**在 `excel.columns` 里定义，否则会报 `引用了 'xxx'，但 excel.columns 未定义该列` 的错误。

### validation 段详解

用于校验工资表数据是否正确，校验结果会写入 Excel 的「验证结果」sheet。

| 配置项 | 必填 | 说明 |
|--------|------|------|
| `enabled` | 否，默认 `false` | 总开关。`true` 启用校验，`false` 关闭 |
| `strict` | 否，默认 `true` | `true`：校验失败时红色报错，禁止提交；`false`：仅黄色警告，仍可提交 |
| `tolerance` | 否，默认 `0.00` | 容差（元）。所有金额先精确到分再比较，差值 ≤ tolerance 算通过 |
| `write_back_sheet` | 否，默认 `true` | 是否把校验结果写入 Excel 新 sheet |
| `write_back_sheet_name` | 否，默认 `"验证结果"` | 写入的 sheet 名称 |
| `column_sum_checks` | 否 | 纵向列加总校验：sum(数据行) == 合计行单元格。只校验列在此名单中的列 |
| `row_formulas` | 否 | 横向公式校验：同时检查合计行和每个数据行是否满足公式 |

#### column_sum_checks 配置

```json
"column_sum_checks": [
  {"column": "deduction_total"},
  {"column": "net_total"},
  {"column": "personal_tax"}
]
```

每条只写一个 `column`（key 名），程序会对该列做：
- 把该列所有数据行相加
- 与合计行该列的数值比较
- 相等（或在 tolerance 内）→ 通过，否则 → 失败

#### row_formulas 配置

```json
"row_formulas": [
  {
    "name": "转款合计 = 扣款合计 + 个人所得税 + 实发工资",
    "lhs": "transfer_total",
    "rhs_plus": ["deduction_total", "personal_tax", "net_total"],
    "rhs_minus": []
  }
]
```

| 字段 | 说明 |
|------|------|
| `name` | 公式名称，显示在校验结果中 |
| `lhs` | 等式左侧的列 key（如 `transfer_total`） |
| `rhs_plus` | 等式右侧相加的列 key 列表 |
| `rhs_minus` | 等式右侧相减的列 key 列表 |

程序会检查：**lhs == sum(rhs_plus) - sum(rhs_minus)**，同时跑合计行和每个数据行。

### 常见问题：校验配置错误

**错误信息**：`validation 配置中 column_sum_checks 引用了 'personal_tax'，但 excel.columns 未定义该列`

**原因**：`validation` 段里的列 key（如 `personal_tax`）在 `excel.columns` 中没有定义。

**解决**：
1. 如果你的工资表**有**「个税」列 → 在 `excel.columns` 里补上：
   ```json
   "personal_tax": {
     "keywords": ["个税", "个人所得税"],
     "label": "个人所得税"
   }
   ```
2. 如果你的工资表**没有**「个税」列 → 从 `validation` 段中移除 `personal_tax` 的引用：
   - `column_sum_checks` 里删掉 `{"column": "personal_tax"}`
   - `row_formulas` 的 `rhs_plus` 里删掉 `"personal_tax"`

**修改后需要重启 streamlit**：配置只在启动时加载一次。

**登录报 400 错误**

1. 确认 `.env` 文件存在于项目根目录（与 `demo_app.py` 同级）
2. 确认 streamlit 从项目根目录启动：`streamlit run demo_app.py`（不是从子目录启动）
3. 重启 streamlit（配置加载只在启动时读取）

**服务启动失败**

1. 检查 `.env` 文件权限：`ls -la .env`
2. 检查日志：`sudo tail -f /home/vod/code/Payroll2DingTalk/streamlit-error.log`
3. 确认虚拟环境路径正确：`which streamlit` 应显示 `.venv/bin/streamlit`

## 已验证的 API 链路

| 步骤 | API | 说明 |
|------|-----|------|
| 0 | `spaces/infos/query` | 每次上传前必须调用，授予临时上传权限 |
| 1 | `topapi/v2/user/getbymobile` | 手机号 → userId |
| 2 | `topapi/v2/user/get` | userId → unionId + deptId |
| 3 | Storage v1.0 `uploadInfos/query` | 获取 OSS 上传凭证 |
| 4 | PUT OSS | 直传文件二进制 |
| 5 | Storage v1.0 `commit` | 提交确认 → fileId |
| 6 | `processInstances` | 创建审批实例 |

## 文档

- `.omo/drafts/dingtalk-api-reference.md` — 钉钉 API 参考（含所有踩坑记录）
- `.omo/drafts/dingtalk-upload-flow.md` — 上传附件完整流程详解

## 技术栈

- Python 3.12
- Streamlit（Web UI）
- requests（HTTP）
- openpyxl（Excel 解析）
- python-dotenv（配置管理）

## 作者

wanghannew1
