---
name: data-consolidation-and-filtering
description: "Use when 批量表格汇总成宽表或列表筛选去重，或要按模板批量出监测曲线图并交付四件套。含口径·抽样·校验·图表套版式。"
version: 1.3.0
license: MIT
metadata:
  tags: [Data, Excel, ETL, Aggregation, Filtering, Sampling, Validation, Charts, Monitoring]
---

# 数据汇总 / 处理 / 筛选 方法论

跨领域通用的四件事：**长表→宽表汇总**、**字段清洗与口径统一**、**条件筛选与去重**、
**抽样与校验**。适用于把一个目录的同构报表合成一张总表，也适用于从接口/列表页里
筛出关心的条目并持续跟踪，以及把 Excel 内嵌曲线图批量套版式、出「曲线图 + 三张表」交付包。

## 0. 执行协议（agent 必读）

这一节是硬约束：**先选定路线，再预检，再按命令跑，最后必须过验收**。
操作细节、输入样例、失败处理见 [`references/agent-runbook.md`](references/agent-runbook.md)，
路线 C 开工前必须读一遍。

### 0.1 定位技能目录（不要依赖 cwd）

`skill_view` 返回的 `_source_path` 去掉 `/SKILL.md` 就是技能目录。后续命令一律用绝对路径：

```bash
SKILL_DIR=/opt/data/skills/productivity/data-consolidation-and-filtering   # 换成实际路径
```

### 0.2 选路线

| 用户要什么 | 路线 | 关键脚本 | 必交产物 |
|---|---|---|---|
| 一批同构报表 → 一张「实体×时间」总表 | **A 汇总** | `long_to_wide.py`（零依赖） | 宽表 + 明细 + 说明页 |
| 从列表/接口按条件筛条目、持续跟踪 | **B 筛选** | 按 §4 自写筛选 + 状态去重 | 命中清单 + 增量状态 |
| 按模板出曲线图、四件套、曲线图+三张表 | **C 交付** | `xlsx_chart_uniform.py` → `xlsx_charts_to_png.py` → `make_delivery.py` | 5 个文件（§9） |

路线 C 需要**原始 Excel 内嵌图表工作簿**；只有数据表时走 A，不要假装能套版式。

### 0.3 预检（阻塞项非 0 不许开工）

```bash
python3 "$SKILL_DIR/scripts/preflight.py" --route C --summary 汇总.xlsx --workbook 原图.xlsx
```

`--summary` 会实际打开工作簿核对表头，`--workbook` 会数内嵌图表数量；正式交付建议带上。

- 退出码 0：开工。WARN 可接受（如无中文字体，图内中文可能成方块）。
- 退出码 1：按输出 `fix` 装依赖后重跑。典型：
  ```bash
  python3 -m venv .venv && .venv/bin/pip install -i https://pypi.tuna.tsinghua.edu.cn/simple \
      openpyxl pillow matplotlib
  ```
  装完统一用 `.venv/bin/python` 跑后续脚本。

### 0.4 路线 C 的固定五步

```bash
PROJ="<项目名>"; WORK="output/<项目>-curves"; PY=.venv/bin/python

"$PY" "$SKILL_DIR/scripts/xlsx_chart_uniform.py" --init 原图.xlsx > "$WORK/chart.json"
# chart.json 可直接用（每表一组、ref=第1张）；跨表统一时才编辑 groups/ref
# 退出码 2 = 图数据被改动，禁止继续
"$PY" "$SKILL_DIR/scripts/xlsx_chart_uniform.py" 原图.xlsx "$WORK/统一版式.xlsx" -c "$WORK/chart.json"
# auto 自动区分时程/测斜，测斜落到 png/incl/
"$PY" "$SKILL_DIR/scripts/xlsx_charts_to_png.py" "$WORK/统一版式.xlsx" "$WORK/png" --dpi 220
"$PY" "$SKILL_DIR/scripts/make_delivery.py" --summary 汇总.xlsx \
    --groups "$WORK/png/_groups.json" --charts "$WORK/png" --sel "$WORK/sel.json" \
    --title "$PROJ" --out "$WORK/deliverables"
"$PY" "$SKILL_DIR/scripts/verify_delivery.py" --dir "$WORK/deliverables" --title "$PROJ" \
    --expect-incl <孔数> --strict
```

- **不要把 `--incl` 指向空目录**：不传时自动用 `<charts>/incl`；旧模板要求文件名含「全过程」，
  新版导出脚本已自动补。
- 第 2 步退出码 2 = 没识别出时程图 → 用 `--time-sheets '图1*'` 显式指定工作表。
- 第 5 步退出码 2 = 有 FAIL，**不得交付**；按 FAIL 行修复后重跑。

### 0.5 输入与输出契约（摘要）

汇总工作簿必须满足：测项表 `第2行 A=点号, B=YYYY-MM-DD`；测斜表 `A列=深度` 且表名以
`测斜` 开头。完整结构、`_groups.json` / `sel.json` 的 JSON 契约见
[`references/agent-runbook.md`](references/agent-runbook.md) §3–§5。

交付目录必须有且仅有 5 类文件：

```
<项目>_数据汇总表.xlsx            <项目>_曲线图.pdf
<项目>_筛选后数据汇总表.xlsx      <项目>_曲线图_PNG.zip
<项目>_各点最大累计值汇总表.xlsx
```

### 0.6 完成报告（先结果后做法）

必须包含：交付件与页/张数、测项/测点/孔数与日期区间、关键极值（含点号/日期/深度）、
`verify_delivery` 的 FAIL/WARN 计数、源数据异常。模板见 runbook §10。

### 0.7 硬约束

- 缺数据源、口径冲突、项目名不确定 → **停下来问**，不猜、不拼、不自己起名。
- 没测 ≠ 测得 0：空就是空，不补 0、不插值。
- 重复主键不静默覆盖，冲突必须留痕。
- 不改验收判据来「让它通过」；不用截图/占位图冒充 PNG。
- 只有 verify 通过（FAIL 0）才算完成，不能只交「脚本跑完了」。

---

## 1. 先定口径，再动代码

开工前必须问清（问错了后面全部返工）：

| 问题 | 为什么 |
|---|---|
| 单元格填**实测值 / 本次变化量 / 累计变化量** | 同一张表三种口径都常见，混用会得出相反的结论 |
| 谁是**行键**、谁是**列键** | 「按点号和日期汇总」= 行点号 × 列日期，不是反过来 |
| 缺失是**真的没有**还是**没测** | 决定留空、补 0 还是标注断档 |
| 重复日期怎么处理 | 后写覆盖先写、还是两条都留 |

## 2. 汇总：长表 → 宽表

```
行 = 实体（点号 / 孔号 / 门店 / 型号）
列 = 时间（日期 / 期次）
值 = 度量（累计变化量 / 销量 / 库存）
```

必须同时输出三类东西：

1. **宽表**：一行一个实体，一列一个日期，冻结首行首列。
2. **长表（明细）**：`实体, 日期, 值, 来源文件` —— 出问题时唯一能追溯的东西。
3. **说明页**：口径、规则、异常清单、完整性检查。

`scripts/long_to_wide.py` 是最小可用实现（零依赖）：

```bash
python3 "$SKILL_DIR/scripts/long_to_wide.py" long.csv wide.csv \
    --entity 点号 --date 日期 --value 累计变化量 --precision 3
# → wide.csv（实体×日期） + wide_conflicts.csv（覆盖留痕）
```

自己写的时候，冲突必须留痕，不能静默覆盖：

```python
wide = {entity: {date: value}}
for f in files:
    for entity, date, value in parse(f):
        if date in wide[entity]:
            conflicts.append((entity, date, wide[entity][date], value, f))
        wide[entity][date] = value
```

分组（按项目/按首次出现日期分组）时可以在组间插标题行，但**读回时必须以
「首列非数值即 break」收尾**，否则下方的注释块会被当成数据，画图会出现贯穿全图的直线。

## 3. 处理：清洗与口径统一

- **主键归一**：去首尾空格、换行、全角空格；保留大小写差异（`ZQT1` 与 `zqt1` 先确认是不是同一个）。
- **同名不同物**：`'建筑物倾斜'` 与 `'建筑物倾斜  '`（尾随空格）是**两张不同的表**，
  点号体系都可能不同。匹配顺序硬规则：精确 → 空格特征相同的 stripped →
  仅当请求名不带空格时才允许 stripped 兜底。
- **日期取值优先级**：文件名日期 > 表内日期。理由：报表里常有「表没更新、日期还停在上期」的僵尸 sheet。
- **空列降级**：累计列为空 → 用变化量按日期逐点累加，并标注「由变化量累加」。
- **隐藏内容**：Excel 用 `wb._sheet_visibility`（与 `sheet_names()` 等长，0=可见）过滤；
  被过滤掉的清单要单独出一页，避免被误解成漏做。
- **僵尸数据检测**：某实体不同取值数 = 1 且样本数很多 → 标记「疑似未更新」，不要当成果交付。
- **重复实体检测**：两条序列重叠日期 ≥80% 且其中 ≥95% 数值相同 → 判为同一实体的两套记录
  （人工表 vs 自动化表），写进备注列，**不要静默合并**。

## 4. 筛选

### 4.1 关键词门控

用「地区 × 主题 × 动作」三个条件**全部命中**才留，避免召回噪音：

```python
if '某区' not in title:                                  continue
if '建设工程' not in title and '工程师' not in title:    continue
```

- 在**自家站点**（地区名不会出现在标题里）要关掉地区条件，否则全被过滤光。
- 过滤掉「说明 / 结论 / 合计 / 测点编号」等伪实体名。
- 接口带 `keyword` 参数时先验证语义：**不少政务/人才站的关键词是精确匹配**，
  传半个词返回 0 条、传空才返回全量 → 正确做法是拉全量再本地 `in` 判断。

### 4.2 滚动窗口陷阱（公告 / 榜单类）

平台列表常常只保留最近 N 条（例：某职称公示列表约 7 天、150 条），过期条目掉出列表，
但详情页长期可用。因此：

- 必须**拉全量（分页到 total 为止）再本地过滤**，不能靠接口关键词；
- 轮询周期 ≤ 窗口长度 / 3（7 天窗口 → 2 小时足够，别用每天一次）；
- 再挂一个**持久列表**作兜底源（结果公布页、归档页），防止长时间停机漏事件。

### 4.3 数值筛选

- 「最大值」先明确是**绝对值最大**还是**代数最大**：位移/沉降要看绝对值，保留符号输出。
- 同一数值多期重复出现时取**最早**一期（首次达到）。
- 单位不同不能放同一列比较（mm / ‰ / kN），分表或在单位列标明。
- 精度按单位给：mm 3 位、‰ 5 位（0.0024 这种值按 3 位就抹平了）、kN 2 位。

## 5. 抽样：代表点怎么选

最容易做错的一步：**按序号等距取点**。

| 目的 | 选法 | 结果特征 |
|---|---|---|
| 画图展示（曲线要铺得开） | 按**度量值分位等距**取目标值，各取最近实测点 | 第 1 条接近最小值、第 k 条接近最大值，中间均匀 |
| 趋势判断（全过程分析） | 按**时间轴等距** 0/20/…/100% 吸附到实测日 | 前段密、后段稀，但覆盖全程 |
| ~~按序号等距~~ | **避免** | 采样前密后疏时，第 1 条≈0、第 2 条已到峰值、其余挤成一团 |

`scripts/pick_representatives.py` 两种模式都实现了：

```python
from pick_representatives import pick
dates, amp = pick({date: [values...]}, k=6, mode="metric")   # mode="time" 切换
```

跑 `python3 "$SKILL_DIR/examples/demo.py"` 看差别：

| 选法 | 6 个代表点（日期 → 特征值 mm） |
|---|---|
| 按度量分位等距 | 07-30→0.00 ・ 08-03→5.20 ・ 08-07→10.40 ・ 08-11→15.60 ・ 08-16→22.10 ・ 11-26→27.00 |
| 按时间等距 | 07-30→0.00 ・ 08-22→26.03 ・ 09-15→26.27 ・ 10-09→26.52 ・ 11-02→26.76 ・ 11-26→27.00 |

补两条：用户说「随机取六天」时，实际要的是**能反映全过程**的六天 ——
用分位/等距取并说明理由，不要真随机。选完把「序号 / 日期 / 特征值 / 量程占比」
四列附在表下方备查。

## 6. 增量与去重（监控类任务）

```python
state = load(STATE)                       # {"源:id": {"t":..., "u":..., "ts":...}}
if not state:                             # 首次运行只写基线，绝不推送历史
    write(state); return
new = [it for it in items if it["key"] not in state]
if not new: return                        # 静默 = 不打扰
```

- 去重键 = `源 + 平台ID`，不要用标题（标题会带日期前缀、会改字）。
- 首次运行必须先**手工跑两遍**建基线，再挂定时；否则第一条实时事件会被当成基线吞掉。
- 状态文件按 `ts` 清理（保留 180~400 天），否则无限增长。
- 无新增时 stdout 为空 —— 配合「空输出不发送」的调度器即可做到零打扰。

## 7. 校验（交出去之前必做）

1. **抽样回源**：随机 2~3 个日期 × 若干实体，逐格与源文件比对，打印 `ok/bad` 计数。
2. **完整性**：日期数是否等于文件数、有无重复日期、有无整列空、有无断档。
3. **量级合理性**：最大值、均值、条数与直觉是否一致（突然大 100 倍通常是单位或日期串列）。
4. **读回验证**：把输出重新打开读一遍，确认注释块没被当数据、合并单元格没串行。

路线 C 用 `scripts/verify_delivery.py` 做机器验收（见 §0.4 第 5 步），人工抽检 5 条见
runbook §8。

## 8. 图表批量统一版式（Excel 内嵌图表）

一堆图要「照着我调好的第一张改」，别手点。`.xlsx` 是 zip，图表在 `xl/charts/chartN.xml`，
改 XML 比逐个开图快两个数量级，而且数据可以做到零改动。

```bash
python3 "$SKILL_DIR/scripts/xlsx_chart_uniform.py" --init 曲线图.xlsx > chart.json
python3 "$SKILL_DIR/scripts/xlsx_chart_uniform.py" 曲线图.xlsx 输出.xlsx -c chart.json
python3 "$SKILL_DIR/scripts/xlsx_charts_to_png.py" 输出.xlsx out/png \
    --dpi 220 --time-size 186x56 --incl-size 90x210
```

**深度剖面别固定条数**：测斜这类「深度 × 日期」矩阵，用
`scripts/pick_profile_curves.py` 按线性插值重构误差自动定张数（首/末/峰值必选，
相邻剖面差 <1.5 mm 视为重复），实测 10 个孔得到 5~9 条，都比硬取 6 条更贴合发展过程。
出图时峰值曲线加粗（一律实线，不用虚线），9 色 + 空心异形标记，首/峰/末打标记，8~9 条也不糊。

**只抄外观，不抄身份**：绘图区几何 / 图例块 / 轴刻度 / 标题格式 ← 基准图；
`axId`、轴标题文字、图名文字、数据 ← 本图。复制图例时必须删 `<c:legendEntry>`。

**量程保护**：本图原量程 < 基准的 40% 就保留本图刻度，否则曲线被拉平（轴力图常见）。
其余按基准 `majorUnit` 逐格外扩包住数据。

**按图分块**：先按测点前缀分类、类内自然排序、每 5 点一张图；`_groups.json` 给出每图测点，
用它把「图 + 绘图数据」逐块嵌进工作簿（每块上图为 186×56 mm、下表日期为行、测点为列）。

**测点排序用自然序**：`DL1、DL2…DL10`，别按字符串排（会成 `DL1、DL10、DL11…DL2`）。
`xlsx_charts_to_png.py` 会按自然序重排每组曲线再分图，图例、表格、最大值清单三处一致。

**图名清洗**：先删 `1/4` 这类序号，再删结尾的 `（mm）`——单位常躲在序号前面。
说明页的清单在 `sharedStrings.xml`，单位可能在字符串中间，用非锚定正则。

**必须校验**：逐图比对输出与源的 `<c:numCache>`（脚本退出码 2 = 有图被改动），
再核对图名集合、组内样式唯一、PNG 抽检。完整模板、规格尺寸、踩坑清单见
[`references/chart-template.md`](references/chart-template.md)。

## 9. 一条龙交付：曲线图 + 三张表

监测类项目最终交 5 个文件（俗称四件套）：曲线图 PDF、数据汇总表、筛选后数据汇总表、
各点最大累计值汇总表、PNG 包。全流程三个脚本，中间只认 `_groups.json`：

```bash
python3 "$SKILL_DIR/scripts/xlsx_chart_uniform.py" 原图.xlsx 统一版式.xlsx -c chart.json
python3 "$SKILL_DIR/scripts/xlsx_charts_to_png.py" 统一版式.xlsx out/png --dpi 220
python3 "$SKILL_DIR/scripts/make_delivery.py" --summary 汇总.xlsx \
    --groups out/png/_groups.json --charts out/png --incl out/png/incl \
    --sel 选点说明.json --title "<项目>" --out deliverables
python3 "$SKILL_DIR/scripts/verify_delivery.py" --dir deliverables --title "<项目>" --strict
```

`make_delivery.py` 产出：`_数据汇总表.xlsx`（全量）、`_筛选后数据汇总表.xlsx`（按图分块：上图下表）、
`_各点最大累计值汇总表.xlsx`（含出现日期/深度）、`_曲线图.pdf`（封面+4 张/页+1 页/孔）、`_曲线图_PNG.zip`。

**节奏**：日报打包并发出去之后不要停下等指令——立刻接着解析数据、汇总、出四件套，
把结果和关键极值一起发回；只有缺数据源或口径冲突时才回头问。

完整规格（轴规则、配色标记、分块命名、校验清单、10 条坑）见
[`references/delivery-workflow.md`](references/delivery-workflow.md)；
执行步骤与失败矩阵见 [`references/agent-runbook.md`](references/agent-runbook.md)；
图表版式的 XML 级细节见 [`references/chart-template.md`](references/chart-template.md)。

## 10. 坑

- **容器里常常没有 `unzip` / `libreoffice`**：解包用 `zipfile`，读老 `.xls`
  （OLE2，文件头 `D0 CF 11 E0`）用 `xlrd`，openpyxl 读不了。
- **`/tmp/inspect.py` 会遮蔽标准库 `inspect`**，导致 `xlrd` 导入报
  partially initialized module。脚本名加前缀。
- **中文文件名 scp 会失败**：`ssh host "cd dir && tar cf - ." | tar xf - -C dest`。
- **matplotlib 中文出豆腐块**：无中文字体的环境里图内文字用英文，Excel 表内保持中文；
  新版 `xlsx_charts_to_png.py` 会扫 `~/.fonts`、Noto CJK、文泉驿等，仍找不到时用 `--font` 指定。
- **散点图** `Series(y=深度, x=位移)` 且 `y_axis.scaling.orientation="maxMin"`
  才能让深度向下增；写反会得到一条贯穿全图的对角斜线。
- **不要给缺失补 0**：`没测` 和 `测得 0` 在趋势判断上是两回事。
- **图表 XML 的 `chartN.xml` 编号不可硬编码**：换一版文件就错位，
  按 `workbook → sheet rels → drawing rels → chart rels` 解析。
- **替换轴后必须重取 `axes(s)`**：每替换一个 `valAx`，后续轴的字符偏移都会变。
- **matplotlib 中文只有 500 字重**：`fontweight="bold"` 会刷 `findfont` 警告，不影响出图。
- **测斜 PNG 不叫「全过程」会被交付脚本漏掉**：`make_delivery` 只认文件名含「全过程」的图；
  新版导出脚本会自动补，手工放图时必须遵守。
