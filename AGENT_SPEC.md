# 订单上传 Agent 需求

## 概述
在 AIWorkAssistant 项目中新增一个 AI Agent："订单自动上传助手"。
该 Agent 监测指定文件夹中的 .doc 文档，解析订单内容，然后通过浏览器自动化登录目标系统并填写销售订单。

## 工作流程
1. 监测文件夹，发现 .doc/.docx 文件
2. 读取文档内容，用 AI 解析成结构化 JSON
3. 打开浏览器（使用 Playwright 或 CefSharp 内嵌浏览器）访问系统 URL
4. 自动填写用户名、密码，弹窗提示用户手动输入验证码
5. 登录成功后跳转到销售订单管理页面
6. 点击新增订单
7. 自动填写基础信息 + 物品信息表格
8. 物品信息：根据物品名称搜索选择，选中后其他信息自动带出，只需填数量和单价

## 登录页结构（Element UI 表单）
- 用户名：input[placeholder="账号"]
- 密码：input[type="password"][placeholder="密码"]  
- 验证码：input[placeholder="验证码"]（需要人工输入）
- 记住我：checkbox
- 登录按钮：button 包含 span "登 录"

## 订单表单字段映射

### 基础信息（from 标签的 for 属性）
| 表单字段 | for属性/标识 | 类型 | 数据来源 |
|---------|-------------|------|---------|
| 销售类型 | saleType | 下拉(其他/煤矿/非煤矿山/非煤/国外) | 配置，默认"煤矿" |
| 销售部门 | deptId | 下拉(很多选项) | 配置，默认"矿用产品销售部" |
| 销售经理 | saleUserId | 下拉(依赖部门) | 配置 |
| 客户名称 | customerId | 搜索下拉 | Word文档"使用单位" |
| 下单日期 | orderTime | 日期选择器 | 当天日期 |
| 合同状态 | 无for | 文本输入(maxlength=10) | 可选，配置 |
| 币种 | moneyType | 下拉(人民币/美元/欧元/港元/日元) | 配置，默认"人民币" |
| 汇率 | rate | 文本输入 | 配置，默认"1" |
| 新老市场 | marketType | 下拉(老市场/新市场) | 配置，默认"老市场" |
| 质保金 | warranty | 数字输入 | 可选 |
| 质保金到期日 | warrantyTime | 日期选择器 | 可选 |
| 备注 | remark | 文本域 | Word文档备注或合同金额信息 |
| 产品线 | productLine | 多选下拉(织布/格栅/其他) | 配置 |
| 发货日期 | deliveryTime | 日期选择器 | Word文档"到货时间"或配置 |
| 订货通知单 | 文件上传 | 上传 | 原始Word文件 |

### 物品信息表格
- 点击"添加物品"按钮添加行
- 在"物品编码"或"物品名称"列搜索选择物品（输入名称关键词匹配）
- 选中物品后，规格型号、计量单位等自动带出
- 需要手动填写：数量、原币含税单价
- 其他金额字段根据数量和单价自动计算

Word文档中的产品及配件清单：
- 产品名称/型号 → 搜索匹配物品
- 数量 → 填入数量列
- 单价 → 填入原币含税单价列

## 可配置项（存储在 SQLite AppSettings 表）
- SystemUrl: 目标系统地址，默认 http://192.168.1.12:8014
- SystemUsername: 登录用户名
- SystemPassword: 登录密码
- DefaultSaleType: 默认销售类型，默认"煤矿"
- DefaultDeptName: 默认销售部门，默认"矿用产品销售部"
- DefaultSaleManager: 默认销售经理
- DefaultMoneyType: 默认币种，默认"人民币"
- DefaultRate: 默认汇率，默认"1"
- DefaultMarketType: 默认新老市场，默认"老市场"
- DefaultProductLine: 默认产品线
- WatchFolder: 监测文件夹路径

## 技术方案
- 浏览器自动化：使用 Microsoft.Playwright 
- Word读取：FreeSpire.Doc（.doc）+ NPOI（.docx）
- AI解析：Anthropic Messages API（已有 ChatService）
- 文件监测：FileSystemWatcher

## 在 AIWorkAssistant 中的集成
- 作为一个 AiAssistant 记录存在数据库中
- 用户选择该助手后，显示专用界面：
  - 监测文件夹路径 + 浏览按钮
  - 启动/停止监测按钮
  - 运行日志
  - 当需要输入验证码时，在界面上显示验证码图片和输入框
