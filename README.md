# AI_coremail_tool
# 技术需求文档

**需求编号**：TRD-2025-001  
**版本**：1.0  
**状态**：评审中  

## 1. 引言

### 1.1 背景与动机
Coremail 邮件系统是公司内部使用的主要邮件系统，通过 https://mail.cffex.com.cn 访问。目前，员工需要手动处理大量邮件，特别是专利相关邮件，包括审查提醒、授权证书、费用发票等。这些邮件处理过程繁琐，容易遗漏重要信息，且缺乏系统化管理。

### 1.2 目标（量化指标）
- 减少 80% 的邮件手动处理时间
- 提高专利审查提醒响应率至 100%
- 实现 95% 的邮件自动分类准确率
- 降低 90% 的专利临期审查遗漏风险
- 提高发票处理效率 70%

### 1.3 范围（含边界排除说明）
**包含**：
- Coremail 邮箱自动连接与同步
- 邮件及附件智能分类与归档
- 专利审查提醒自动化处理
- 发票类附件分类与信息提取
- 与飞书多维表格的集成

**不包含**：
- 邮件撰写与回复功能
- 非专利相关邮件的深度处理
- 邮件加密与解密
- 与其他邮件系统的集成

### 1.4 参考资料
- Coremail API 文档
- 飞书开放平台 API 文档
- 现有 EmailManager 类实现
- 专利审查流程文档

## 2. 需求详情

### 2.1 功能需求

| 功能ID | 描述 | 输入/输出 | 业务规则 |
|--------|------|-----------|----------|
| F-01 | 邮箱自动连接与同步 | 用户名密码→同步状态 | 支持 IMAP/POP3/SMTP 协议，定时自动同步 |
| F-02 | 邮件智能分类 | 邮件内容→分类结果 | 基于内容和发件人进行分类，支持自定义规则 |
| F-03 | 附件自动归档 | 附件→归档路径 | 按附件类型和内容智能归类，支持批量处理 |
| F-04 | 专利审查提醒管理 | 邮件→提醒事项 | 自动提取申请号、期限等信息，按紧急程度分类 |
| F-05 | 发票信息提取与汇总 | 发票附件→结构化数据 | 自动识别发票类型，提取金额、税率等信息 |
| F-06 | 临期提醒自动化 | 期限日期→提醒事件 | 设置多级提醒时间点，支持邮件和飞书通知 |
| F-07 | 飞书多维表格集成 | 专利信息→飞书表格 | 双向同步，支持状态变更触发提醒 |
| F-08 | 专利证书管理 | 证书附件→证书库 | 自动提取专利号、名称等信息，支持检索 |
| F-09 | 软件协会通知分类 | 通知邮件→分类结果 | 自动识别评奖、活动等不同类型通知 |
| F-10 | 费用合计计算 | 多发票→汇总报表 | 支持按专利、时间段等维度统计费用 |

### 2.2 非功能需求

- **性能**：
  - 邮件同步速度≥100封/分钟
  - 附件处理速度≥50MB/分钟
  - Web界面响应时间≤500ms

- **安全**：
  - 邮箱凭证加密存储
  - HTTPS传输
  - 用户权限分级控制

- **可靠性**：
  - 系统可用性≥99.5%
  - 数据备份与恢复机制
  - 同步失败自动重试

- **兼容性**：
  - 支持Chrome/Firefox/Edge最新版
  - 支持Windows/macOS/Linux操作系统
  - 支持移动端访问

- **可扩展性**：
  - 支持多用户并发使用
  - 模块化设计，便于功能扩展
  - 支持API接口调用

## 3. 技术方案

### 3.1 架构图
```
graph TD
    A[Coremail Server] <--> B[Coremail-Connect Application]
    B <--> C[飞书多维表格]
    B --> D[(数据库)]
    B --> E[文件存储]
    B --> F[缓存]
    
    style A fill:#f9f,stroke:#333,stroke-width:2px
    style B fill:#bbf,stroke:#333,stroke-width:2px
    style C fill:#bfb,stroke:#333,stroke-width:2px
    style D fill:#fbb,stroke:#333,stroke-width:2px
    style E fill:#fbf,stroke:#333,stroke-width:2px
    style F fill:#bff,stroke:#333,stroke-width:2px
```
```
graph LR
    A[邮件同步] --> B[邮件分析处理]
    B --> C[数据存储]
    B --> D[信息提取]
    D --> E[临期检测]
    D --> F[附件归档]
    E --> G[飞书通知]
    F --> H[费用计算汇总]
    
    style A fill:#f9f,stroke:#333,stroke-width:2px
    style B fill:#bbf,stroke:#333,stroke-width:2px
    style C fill:#bfb,stroke:#333,stroke-width:2px
    style D fill:#fbb,stroke:#333,stroke-width:2px
    style E fill:#fbf,stroke:#333,stroke-width:2px
    style F fill:#bff,stroke:#333,stroke-width:2px
    style G fill:#ff9,stroke:#333,stroke-width:2px
    style H fill:#f99,stroke:#333,stroke-width:2px
```

### 3.2 技术栈

- **后端**：Python + Flask
- **数据库**：SQLite (本地存储)
- **邮件处理**：imaplib, poplib, smtplib, email
- **文本处理**：NLTK, jieba, re
- **AI分类**：scikit-learn, InternLM API
- **定时任务**：APScheduler
- **前端**：HTML + CSS + JavaScript + Bootstrap
- **API集成**：Requests, 飞书开放平台SDK

### 3.3 数据库设计

**邮件表 (emails)**
| 字段名 | 类型 | 说明 |
|--------|------|------|
| id | INTEGER | 主键 |
| message_id | TEXT | 邮件唯一ID |
| subject | TEXT | 邮件主题 |
| from_addr | TEXT | 发件人 |
| date | DATETIME | 邮件日期 |
| category | TEXT | 分类结果 |
| content_hash | TEXT | 内容哈希值 |
| file_path | TEXT | 本地存储路径 |
| processed | BOOLEAN | 处理状态 |

**专利提醒表 (patent_reminders)**
| 字段名 | 类型 | 说明 |
|--------|------|------|
| id | INTEGER | 主键 |
| email_id | INTEGER | 关联邮件ID |
| application_no | TEXT | 专利申请号 |
| client_no | TEXT | 客户编号 |
| our_no | TEXT | 我方编号 |
| deadline | DATE | 截止日期 |
| urgency_level | TEXT | 紧急程度 |
| completed | BOOLEAN | 完成状态 |
| notify_status | TEXT | 通知状态 |

**发票表 (invoices)**
| 字段名 | 类型 | 说明 |
|--------|------|------|
| id | INTEGER | 主键 |
| email_id | INTEGER | 关联邮件ID |
| invoice_number | TEXT | 发票编号 |
| invoice_type | TEXT | 发票类型 |
| pre_tax_amount | DECIMAL | 税前金额 |
| tax_rate | DECIMAL | 税率 |
| tax_amount | DECIMAL | 税额 |
| total_amount | DECIMAL | 总金额 |
| file_path | TEXT | 文件路径 |

**附件表 (attachments)**
| 字段名 | 类型 | 说明 |
|--------|------|------|
| id | INTEGER | 主键 |
| email_id | INTEGER | 关联邮件ID |
| filename | TEXT | 原始文件名 |
| file_type | TEXT | 文件类型 |
| category | TEXT | 分类 |
| saved_path | TEXT | 保存路径 |
| extracted | BOOLEAN | 信息提取状态 |

**飞书集成表 (feishu_integration)**
| 字段名 | 类型 | 说明 |
|--------|------|------|
| id | INTEGER | 主键 |
| patent_id | INTEGER | 关联专利ID |
| feishu_record_id | TEXT | 飞书记录ID |
| last_sync_time | DATETIME | 最后同步时间 |
| sync_status | TEXT | 同步状态 |

## 4. 验收标准

### 4.1 测试用例

| 测试ID | 测试内容 | 预期结果 |
|--------|----------|----------|
| T-01 | 邮箱连接与同步 | 成功连接邮箱并同步最新邮件 |
| T-02 | 邮件分类准确性 | 分类准确率≥95% |
| T-03 | 专利提醒提取 | 正确提取申请号、期限等信息 |
| T-04 | 发票信息提取 | 准确识别发票类型并提取金额信息 |
| T-05 | 临期提醒功能 | 按设定时间点触发提醒 |
| T-06 | 飞书集成 | 数据成功同步至飞书多维表格 |
| T-07 | 费用合计计算 | 多发票汇总计算结果准确 |
| T-08 | 系统性能 | 满足性能指标要求 |

### 4.2 性能指标

- 邮件同步速度≥100封/分钟
- 附件处理速度≥50MB/分钟
- Web界面响应时间≤500ms
- 系统CPU占用率≤30%
- 内存占用≤500MB

## 5. 项目计划

| 阶段 | 起止日期 | 负责人 |
|------|----------|--------|
| 需求分析与设计 | 2025-07-01~2025-07-10 | 项目经理 |
| 邮箱连接与同步模块 | 2025-07-11~2025-07-20 | 后端开发 |
| 邮件分类与附件归档 | 2025-07-21~2025-08-05 | 算法工程师 |
| 专利提醒与发票处理 | 2025-08-06~2025-08-20 | 后端开发 |
| 飞书集成与通知 | 2025-08-21~2025-09-05 | 集成工程师 |
| Web界面开发 | 2025-08-06~2025-09-10 | 前端开发 |
| 系统测试 | 2025-09-11~2025-09-25 | 测试工程师 |
| 用户验收测试 | 2025-09-26~2025-10-10 | 项目经理 |
| 部署上线 | 2025-10-11~2025-10-15 | 运维工程师 |

## 6. 附录

### 6.1 风险分析表

| 风险ID | 风险描述 | 可能性 | 影响 | 缓解措施 |
|--------|----------|--------|------|----------|
| R-01 | 邮箱服务器不稳定 | 中 | 高 | 实现断点续传，失败重试机制 |
| R-02 | 邮件格式多样导致解析错误 | 高 | 中 | 增强解析算法，添加异常处理 |
| R-03 | 飞书API变更 | 低 | 高 | 模块化设计，快速适配新API |
| R-04 | 性能瓶颈 | 中 | 中 | 优化算法，实现增量同步 |
| R-05 | 数据安全风险 | 低 | 高 | 加密存储，权限控制，日志审计 |

### 6.2 术语表

| 术语 | 定义 |
|------|------|
| Coremail | 公司使用的邮件系统 |
| IMAP/POP3 | 邮件接收协议 |
| SMTP | 邮件发送协议 |
| 专利申请号 | 专利的唯一标识号 |
| 临期提醒 | 专利审查期限临近的提醒 |
| 飞书多维表格 | 飞书提供的在线协作表格工具 |
| 发票类型 | 包括官方票据、代理票据、代理XML文件等 |
