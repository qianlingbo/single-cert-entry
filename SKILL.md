# 单证录入技能

> 将船舶 IMO Crew List + Port of Call 原始文件，一键转换为海事局标准录入格式（**.xlsm**）

## 版本信息

| 项目 | 内容 |
|------|------|
| 技能版本 | **v2.0.0** (.xlsm 17-sheet 模板) |
| 最后更新 | 2026-06-28 |
| 状态 | 🟢 生产就绪 |
| Python | 3.8+ |
| 依赖 | `openpyxl` |

---

## 目录结构

```
单证录入工作区/
├── SKILL.md                          # 本文档
├── README.md                         # 使用说明
├── .gitignore                        # 排除 input/output
├── templates/
│   ├── 标准格式.xlsm                 # ⭐ v2 模板 (17 sheet, 含完整海事单证)
│   └── 单证录入标准格式.xlsx         # v1 模板 (3 sheet, 旧版)
├── scripts/
│   ├── 单证录入核心_v2.py            # ⭐ v2 主脚本 (.xlsm 17 sheet)
│   └── 单证录入核心.py               # v1 脚本 (3 sheet, 旧版)
├── references/                       # v1 参数映射（v2 内置映射已够用）
│   ├── nationality_map.json          # 248条 国籍代码
│   ├── duty_map.json                 # 12条  职务代码
│   └── port_map.json                 # 1956条 港口代码
├── input/                            # 放原始文件（不入库）
└── output/                           # 生成的文件（不入库）
```

---

## 🚀 快速开始

```bash
# 1. 进入工作区
cd ~/单证录入工作区

# 2. 安装依赖
pip install openpyxl

# 3a. v2 主脚本（推荐，输出 .xlsm 17 sheet）
python3 scripts/单证录入核心_v2.py

# 3b. v1 旧脚本（输出 .xlsx 3 sheet）
python3 scripts/单证录入核心.py input/crew_list.xlsx input/port_of_call.xlsx 2025航次报告
```

---

## 📊 v2 输出 Sheet 清单（17 个）

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

---

## 📊 单证录入规则

### 1. 船员名单（船上非旅客人员清单）

| 字段 | 规则 |
|------|------|
| **序号** | 按顺序递增，从1开始 |
| **姓名** | 中国船员 = 中文；外国船员 = 大写英文字母 |
| **性别** | `1-男` / `2-女` |
| **船员职务** | 英文职务 → 代码（51-船长 / 52-大副 / 53-二副 / 54-三副 / 55-值班水手 / 56-高级值班水手 / 61-轮机长 / 62-大管轮 / 63-二管轮 / 65-值班机工），找不到=55 |
| **船员国籍** | `CN-中国` / `VN-越南` / `MM-缅甸` / `ID-印度尼西亚` 等 |
| **出生日期** | `YYYYMMDD` 格式（如 `19680625`） |
| **出生地点** | **取国籍的中文名**（如 VU VAN TRONG → 越南；姜文敏 → 中国） |
| **证件类型** | 中国船员 → `17-海员证`；外国船员 → `14-普通护照` |
| **证件号码** | 中国船员 → 海员证号码；外国船员 → 护照号码 |
| **是否申请登陆** | 留空 |
| **适任证书编号** | 留空 |
| **适任证书有效期至** | 留空 |
| **证件检查地点** | 留空 |
| **登船日期** | `YYYYMMDD` 格式 |
| **登船口岸** | 港口代码，找不到保持空白 |

### 2. 船上非旅客人员物品清单

| 字段 | 规则 |
|------|------|
| **序号** | 同船员序号 |
| **证件类型** | 参照船员名单（17-海员证 / 14-普通护照） |
| **证件号码** | 参照船员名单 |
| **物品类型** | `0100`（计算机） |
| **物品名称** | `计算机` |
| **物品数量** | `1` |

### 3. 海事船岸活动信息（Port of Call）

| 字段 | 规则 |
|------|------|
| **序号** | 按顺序递增（自动重排） |
| **进港时间** | `YYYY/MM/DD HH:MM:SS`，**时间随机取 00:00–11:59** |
| **离港时间** | `YYYY/MM/DD HH:MM:SS`，**时间随机取 12:00–23:59** |
| **国家/地区名称** | 两字母代码（`CN` / `ID` / `UN` 公海） |
| **船舶保安等级** | `1-1级` |
| **特别/附加保安设施** | 留空 |
| **停靠港口** | 港口代码（CNZOS / CNZJG / CNTXG 等），找不到保持空白 |
| **港口保安等级** | `1-1级` |

### 4. 前十港信息

字段同海事船岸活动（不含保安等级列）。

---

## 🔧 v2 核心映射表

### 国籍代码

| 英文 | 中文 | 代码 |
|------|------|------|
| CHINESE | 中国 | CN |
| VIETNAM | 越南 | VN |
| MYANMAR | 缅甸 | MM |
| INDONESIA | 印度尼西亚 | ID |
| PANAMA | 巴拿马 | PA |
| INDIA | 印度 | IN |
| PHILIPPINES | 菲律宾 | PH |

### 职务代码

| 英文 | 代码 | 中文 |
|------|------|------|
| MASTER | 51 | 船长 |
| C/O | 52 | 大副 |
| 2/O | 53 | 二副 |
| 3/O | 54 | 三副 |
| OS / OLR / AB / C/CK | 55 | 值班水手 |
| BOSUN | 56 | 高级值班水手 |
| C/E | 61 | 轮机长 |
| 2/E | 62 | 大管轮 |
| 3/E | 63 | 二管轮 |
| ETR / FTR | 65 | 值班机工 |

### 港口代码（部分）

| 英文 | 代码 |
|------|------|
| ZHOUSHAN | CNZOS |
| ZHANGJIAGANG | CNZJG |
| CHANGZHOU | CNCZX |
| TIANJIN | CNTXG |
| LANSHAN | CNLSN |
| BAYUQUAN | CNBYQ |
| CAOFEIDIAN | CNCFD |
| LIANYUNGANG | CNLYG |
| NINGDE | CNNDS |
| MOROWALI | IDMOR |
| POMALAA | IDPUM |
| OBI ISLAND | IDOBI |
| OPEN SEA | OPSEA |

**未在标准参数表的港口**（保持空白）：
- CHENJIAGANG（陈家港，CN）
- KENDARI（肯达里，ID）

完整 1958 条见 `templates/标准格式.xlsm` 的「参数」sheet。

---

## 🔧 核心函数（v2 内部）

### `parse_crew_xlsx(path)` → 船员列表
解析 IMO Crew List 的 `ARR` sheet（每两行一条记录：英文行 + 中文行）。

### `parse_poc_xlsx(path)` → 港口列表
解析 Port of Call 的 `Sheet1`（从第 4 行开始）。

### `normalize_date(s)` → datetime
支持 `2026/07/02 09:18:00`、空格不规则、月份内空格等情况（正则去空+提取）。

### `map_nationality(nat_en)` → `CN/VN/MM/...`
英文/中文 → 国籍代码。

### `map_birth_place(nat_en)` → `中国/越南/缅甸/...`
**出生地点强制取国籍的中文**（按规则）。

### `map_port_code(port_en)` → `CNZOS/...` 或 None
港口 → 代码；找不到返回 `None`（保留空白）。

### `map_signon_port_code(port_en)` → `CNNDS/...` 或 None
登船地点 → 登船口岸代码。

---

## 📁 输入文件格式

### IMO Crew List（Excel）

典型表头含关键词：`Name of Ship` + `Port of Arr` + `Date of Arrival`（第 1-5 行）
数据行 R7 是列名，R9 起是数据（每两行：英文行 + 中文行）

列布局：
```
A=序号  B=Family name  C=Sex  D=Rank  E=Date of Birth  F=Nationality
G=Place & Date Signed on  H=Seaman's Book No.  I=Expiry (seaman)  J=Passport No.
K=Expiry (passport)
（中文行：B=中文名  E=出生地点  G=登船地点）
```

### Port of Call List（Excel）

典型表头 R3，列布局：
```
A=No.  B=Purpose  C=Port  D=Country
E=Arrival Time  F=Departure Time  G=Ship Security Level  H=Port Security Level
```

---

## 📌 已知局限

1. **PDF 支持**：v2 暂未实现 PDF 解析
2. **未找到的港口**：保持空白（用户可手动填写）
3. **物品清单**：固定为每人 1 件计算机（按 SKILL.md 默认规则）
4. **随机时间**：v2 严格使用 `00-11` / `12-23`，可能与原始时间差异较大
5. **序号重排**：原 Port of Call 序号重复时按出现顺序重排

---

## 🛠️ 调试技巧

```bash
# v2 主脚本帮助
python3 scripts/单证录入核心_v2.py --help

# 指定输入输出
python3 scripts/单证录入核心_v2.py \
    input/CREW_LIST.xlsx \
    input/PORT_OF_CALL.xlsx \
    output/ZHIDA2_单证录入.xlsm
```

---

## 📄 变更历史

- **v2.0.0 (2026-06-28)** — 基于 `.xlsm` 17-sheet 模板重构
  - 改用 `templates/标准格式.xlsm` 作为基础模板
  - 自动填充 4 个 sheet：船员名单 / 物品清单 / 海事船岸活动 / 前十港
  - 出生地点强制使用国籍中文
  - 时间规则严格 00-11 / 12-23
- **v1.0.0 (2026-04-10)** — 初始版本，3 sheet 输出
