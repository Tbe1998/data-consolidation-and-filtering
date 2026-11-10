# 数据汇总 / 处理 / 筛选方法论 v1.3

把一堆同构报表合成一张「实体 × 时间」宽表，从长列表里按条件筛选去重，再挑出能铺开的
代表数据，最后把 Excel 里几十张内嵌图表批量统一版式、出「曲线图 + 三张表」交付包。
这套套路反复出现在报表合并、监测日报汇总、公告监控、榜单抓取、曲线图返工里。
这里把它写成**可执行、可验收的 skill + 脚本**，而不是一篇泛泛而谈的规范。

```text
行 = 实体（点号 / 孔号 / 门店 / 型号）
列 = 时间（日期 / 期次）
值 = 度量（累计变化量 / 销量 / 库存）
```

## 目录

| 路径 | 内容 |
|---|---|
| `SKILL.md` | 方法论 + **执行协议**：口径 → 汇总 → 清洗 → 筛选 → 抽样 → 去重 → 校验 → 交付 |
| `references/agent-runbook.md` | **给执行 agent 的操作手册**：输入契约、五步命令、输出契约、失败矩阵、报告模板 |
| `scripts/preflight.py` | 环境自检：Python/依赖/中文字体/脚本完整性，退出码 1 = 有阻塞 |
| `scripts/long_to_wide.py` | 长表 CSV → 宽表 CSV，冲突留痕，零第三方依赖 |
| `scripts/pick_representatives.py` | 代表点抽样：按度量分位等距 / 按时间等距 |
| `scripts/xlsx_chart_uniform.py` | 批量统一 Excel 内嵌图表版式，数据零改动（零依赖） |
| `scripts/xlsx_charts_to_png.py` | 自动区分时程/测斜并导出 PNG（需 matplotlib），产出 `_groups.json` |
| `scripts/pick_profile_curves.py` | 深度剖面曲线数量自适应：插值重构误差 + 重复剔除（零依赖） |
| `scripts/make_delivery.py` | 一条龙交付：曲线 PDF + 3 张 Excel + PNG 包（需 openpyxl/Pillow） |
| `scripts/verify_delivery.py` | **交付验收**：页数/张数/嵌图数/行数/分组/图名，FAIL 非 0 不得交付 |
| `references/chart-template.md` | 监测曲线模板：规格、谁抄谁、量程保护、校验清单、踩坑 |
| `references/delivery-workflow.md` | 完整交付工作流：轴规则、配色标记、分类分块、校验清单、10 条坑 |
| `examples/chart_uniform.example.json` | 套版式配置示例（分组 / 基准图 / 图例改日期） |
| `examples/sample_long.csv` | 带冲突、空值、非法值的样例长表 |
| `examples/demo.py` | 两种抽样模式的效果对比 |
| `tests/` | 单测 + 端到端冒烟（`python3 -m unittest discover -s tests -v`） |
| `.github/workflows/ci.yml` | Python 3.8 / 3.11 / 3.12 上跑编译、单测与冒烟 |

## 快速开始

路线 A/B 只需要 Python 3.8+，零第三方依赖；
路线 C（曲线图四件套）需要 `openpyxl`、`Pillow`、`matplotlib`。

```bash
# 0) 先自检（任何路线都建议跑）
python3 scripts/preflight.py --route C
# 缺依赖时：
python3 -m venv .venv
.venv/bin/pip install -i https://pypi.tuna.tsinghua.edu.cn/simple openpyxl pillow matplotlib
# 之后用 .venv/bin/python 跑脚本

# 1) 长表 → 宽表（自动生成 xxx_conflicts.csv 记录覆盖行为）
python3 scripts/long_to_wide.py examples/sample_long.csv out/wide.csv \
    --entity entity --date date --value value --precision 3

# 2) 抽样模式对比
python3 examples/demo.py
```

`scripts/long_to_wide.py` 会做这些事：日期归一（`2024/08/07`、`2024.08.07`、
Excel serial 都认）、主键去空白、空值保持为空（不补 0）、同一实体+日期重复出现时
后写覆盖先写并在 conflicts 文件里留痕、输出按实体与日期排序的确定性 CSV。

### 批量统一 Excel 曲线图版式 → 四件套

```bash
# 1) 看结构（工作表 / 每页几张图），生成配置草稿
python3 scripts/xlsx_chart_uniform.py --init 曲线图.xlsx > chart.json

# 2) 编辑 chart.json：合并同组、ref 指向基准图（示例见 examples/chart_uniform.example.json）

# 3) 套版式（原文件不动；退出码 2 = 有图数据被改动，禁止继续）
python3 scripts/xlsx_chart_uniform.py 曲线图.xlsx 曲线图_统一版式.xlsx -c chart.json

# 4) 重绘 PNG：auto 区分时程/测斜，测斜落到 out/png/incl/ 并带“全过程”
python3 scripts/xlsx_charts_to_png.py 曲线图_统一版式.xlsx out/png --dpi 220

# 5) 四件套（--incl 可省，默认 <charts>/incl）
python3 scripts/make_delivery.py --summary 汇总.xlsx --groups out/png/_groups.json \
    --charts out/png --sel 选点说明.json --title "110二期X3#井" --out deliverables

# 6) 验收：FAIL 0 才算完成
python3 scripts/verify_delivery.py --dir deliverables --title "110二期X3#井" \
    --expect-incl 10 --strict
```

脚本只搬外观（绘图区几何、图例块、轴刻度、标题格式），不搬身份（轴 ID、轴标题文字、
图名文字、数据）。自带 `<c:numCache>` 零改动校验。完整模板与踩坑见
[`references/chart-template.md`](references/chart-template.md)。

### 深度剖面曲线：数量按位移发展自适应

测斜孔有几十上百期时，固定 6 条曲线会漏掉中段转折，全画又糊。用插值重构误差决定张数：

```bash
python3 scripts/pick_profile_curves.py tst2.csv --format text
# 实测 131 期 → 选 9 期  峰值 27.5 mm（2024-11-08）
# 阈值 2.75 mm，最大重构误差 2.43 mm
```

首测/末测/峰值必选，中间只补「线性插值还原不出来的」期次，相邻剖面差 <1.5 mm 视为重复。
110 二期 X3# 井 10 个孔自动得到 5~9 条，峰值曲线加粗标红、峰值后用虚线即可保持清爽。

## 给其他 agent 执行

- 先读 [`references/agent-runbook.md`](references/agent-runbook.md)：输入清单、五步命令、
  输出契约、失败矩阵、完成报告模板都在里面。
- 全流程命令用 `$SKILL_DIR` 绝对路径，**不依赖当前目录**。
- 只有 `verify_delivery.py` 返回 FAIL 0 才允许交付；缺数据/口径冲突时停下来问用户。

## 核心结论（正文里都有展开）

**1. 口径先行。** 实测值 / 本次变化量 / 累计变化量混用，会得出相反的结论；
「按点号和日期汇总」的默认含义是行=点号、列=日期。

**2. 冲突不静默。** 同一主键出现两次就记录，覆盖要有痕迹，否则后面没人能解释数字为什么变了。

**3. 同名不等于同物。** `'建筑物倾斜'` 和 `'建筑物倾斜  '`（尾随空格）是两张不同的表，
点号体系都可能不同。匹配顺序：精确 → 空格特征相同的 stripped → 仅不带空格时才允许 stripped 兜底。

**4. 日期优先级：文件名 > 表内。** 报表里常有「表没更新、日期还停在上期」的僵尸 sheet。

**5. 归属要查重。** 两条序列重叠日期 ≥80% 且其中 ≥95% 数值相同 → 同一实体的两套记录
（人工表 / 自动化表），写进备注，不要静默合并。

**6. 抽样别按序号等距。** 采样前密后疏时，第 1 条≈0、第 2 条已到峰值、其余挤成一团。
画图用**度量分位等距**，趋势分析用**时间等距**：

| 选法 | 6 个代表点（日期 → 特征值 mm） |
|---|---|
| 按度量分位等距 | 07-30→0.00 ・ 08-03→5.20 ・ 08-07→10.40 ・ 08-11→15.60 ・ 08-16→22.10 ・ 11-26→27.00 |
| 按时间等距 | 07-30→0.00 ・ 08-22→26.03 ・ 09-15→26.27 ・ 10-09→26.52 ・ 11-02→26.76 ・ 11-26→27.00 |

**7. 滚动窗口要拉全量。** 平台列表常只留最近 N 条（过期掉出列表、详情页仍可用），
所以必须分页拉到 total 再本地过滤，轮询周期 ≤ 窗口长度 / 3，并挂一个持久列表兜底。

**8. 交付前必须回源校验。** 抽样逐格比对、完整性检查、量级合理性、读回验证，缺一项都可能把错表交出去。

**9. 图表套版式只抄外观，不抄身份。** 绘图区、图例、刻度、字号跟基准图；
轴 ID、轴标题文字、图名、数据留在本图。量程差 2.5 倍以上要保留本图刻度，
否则曲线被拉成一条直线。改完逐图比对 `numCache`，确认一个数据点都没动。

## 作为 skill 安装

`SKILL.md` 带 YAML frontmatter，可直接放进 Hermes / Claude 风格的 skills 目录：

```bash
# Hermes：skills 根目录通常是 $HERMES_HOME/skills（默认 ~/.hermes/skills）；
# 本仓库也可直接放到某个 profile 的 skills/productivity/ 下
git clone https://github.com/Tbe1998/data-consolidation-and-filtering.git \
  "${HERMES_HOME:-$HOME/.hermes}/skills/productivity/data-consolidation-and-filtering"
```

（路径按你的 agent 约定调整；脚本、`references/`、`examples/` 与 `SKILL.md` 同级放好即可。）

## License

MIT
