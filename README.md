# 单证录入工作区

> 船舶单证录入自动化工具 —— 将 IMO Crew List + Port of Call 原始 Excel 文件，一键转换为海事局标准录入格式 (`.xlsm`)。

## 功能

- 🤖 **全自动转换**：原始文件 → 标准格式，无需手动填表
- 📋 **17 个标准 Sheet**：船员名单 / 物品清单 / 前十港 / 海事船岸活动 / 压舱水 / 危险品 / 船舶证书 等
- 🔢 **标准编码**：使用海事局参数字段表（国籍 / 职务 / 港口 / 证件），保证数据合规
- 🇨🇳 **出生地点规则**：统一取**船员国籍的中文**（如 VU VAN TRONG → 越南）
- ⏰ **时间规则**：进港 00-11 点 / 离港 12-23 点（严格按 SKILL.md 规则）

## 脚本版本

| 脚本 | 模板 | 说明 |
|------|------|------|
| `scripts/单证录入核心_v2.py` | `templates/标准格式.xlsm` | **当前推荐** — 输出 17-sheet 完整海事单证 |
| `scripts/单证录入核心.py` | `templates/单证录入标准格式.xlsx` | 旧版本 — 仅输出 3 sheet |

## 环境要求

```bash
# Python 3.8+
python3 --version

# 安装依赖
pip install openpyxl
```

## 使用方法

### 方式一：命令行（v2 推荐）

```bash
cd ~/单证录入工作区

# 默认参数（从 input/ 读取，输出到 output/）
python3 scripts/单证录入核心_v2.py

# 自定义输入输出
python3 scripts/单证录入核心_v2.py \
    input/CREW_LIST.xlsx \
    input/PORT_OF_CALL.xlsx \
    output/ZHIDA2_单证录入.xlsm
```

### 方式二：飞书 / Hermes 对话

直接将 `IMO CREW LIST.xlsx` 和 `PORT OF CALL LIST.xlsx` 拖入飞书/Hermes 对话，AI 自动处理并推送结果文件。

## 输入文件

### IMO CREW LIST.xlsx
| 列 | 含义 | 示例 |
|---|------|------|
| No. | 序号 | 1 |
| Family name | 英文姓名 | JIANG WENMIN |
| (中文行) | 中文姓名 | 姜文敏 |
| Sex | 性别 | MALE |
| Rank | 职务 | MASTER |
| Date of Birth | 出生日期 | 1968-06-25 |
| Nationality | 国籍 | CHINESE |
| Place of Birth | 出生地 | LIAONING |
| Signed on | 登船日期 | 2025-12-17 |
| Place Signed on | 登船地点 | NINGDE |
| Seaman's Book No. | 海员证号 | A90464661 |
| Passport No. | 护照号 | EB9563094 |

### PORT OF CALL.xlsx
| 列 | 含义 | 示例 |
|---|------|------|
| No. | 序号 | 1 |
| Purpose | 目的 | Discharging |
| Port | 港口 | CHENJIAGANG |
| Country | 国家 | CN |
| Arrival Time | 进港时间 | 2026/07/02 09:18:00 |
| Departure Time | 离港时间 | 2026/07/03 09:18:00 |
| Ship Security Level | 船舶保安等级 | 1 |
| Port Security Level | 港口保安等级 | 1 |

## 输出文件

`output/单证录入标准格式.xlsm`，含 **17 个标准 Sheet**：

| # | Sheet | 自动填充 | 说明 |
|---|-------|---------|------|
| 1 | 船上非旅客人员清单 | ✅ | 16 列（国籍中文作为出生地点） |
| 2 | 旅客清单 | — | 无旅客时留空 |
| 3 | 供退物料清单 | — | 待填 |
| 4 | 船上非旅客人员物品清单 | ✅ | 每人 1 件计算机 (0100) |
| 5 | 船用物品清单 | — | 待填 |
| 6 | 前十港信息 | ✅ | 来自 Port of Call |
| 7 | 危险品信息 | — | 待填 |
| 8 | 船舶证书信息 | — | 模板含 12 项标准证书 |
| 9 | 压舱水详细信息 | — | 模板空表 |
| 10 | 海事船岸活动信息 | ✅ | 11 条港口 |
| 11 | 沿海空箱信息 | — | 待填 |
| 12-15 | 压舱水报告单/装载/更换/排放 | — | 待填 |
| 16 | 随船人员清单 | — | 待填 |
| 17 | 参数 | — | 下拉验证数据源 |

## 字段映射规则（详见 SKILL.md）

| 原始字段 | 输出代码 | 备注 |
|----------|----------|------|
| CHINESE / 中国 | CN | 国籍 |
| VIETNAM / 越南 | VN | 国籍 |
| MYANMAR / 缅甸 | MM | 国籍 |
| MASTER | 51 | 职务 |
| C/O | 52 | 职务 |
| 中国船员 | 17 | 海员证 |
| 外国船员 | 14 | 普通护照 |
| ZHOUSHAN | CNZOS | 港口代码 |
| TIANJIN | CNTXG | 港口代码 |
| OPEN SEA | OPSEA | 公海 |

**找不到代码的港口**：保持空白（按 SKILL.md 规则）。

## 目录结构

```
单证录入工作区/
├── SKILL.md                          # 技能详细文档
├── README.md                         # 本文件
├── .gitignore                        # 排除 input/output
├── templates/
│   ├── 标准格式.xlsm                 # ⭐ v2 模板 (17 sheet)
│   └── 单证录入标准格式.xlsx         # v1 模板 (3 sheet)
├── scripts/
│   ├── 单证录入核心_v2.py            # ⭐ v2 主脚本
│   └── 单证录入核心.py               # v1 脚本
├── references/                       # v1 参数映射
│   ├── nationality_map.json          # 国籍代码（248条）
│   ├── duty_map.json                 # 职务代码（12条）
│   └── port_map.json                 # 港口代码（1956条）
├── input/                            # 放原始文件（不入库）
└── output/                           # 生成的文件（不入库）
```

## 工作原理

```
原始文件 → 智能解析（自动找表头/列索引）
        → 标准化映射（国籍/职务/港口/日期）
        → 规则补全（fallback 职务/国家提取）
        → Excel 写入（17 个 Sheet）
        → 输出 .xlsm 文件
```

## 技术栈

- **Python 3.8+**
- **openpyxl**：Excel 读写

## 许可证 & 作者

MIT License
Built with [Hermes](https://github.com/) + Claude MiniMax-M2.5
