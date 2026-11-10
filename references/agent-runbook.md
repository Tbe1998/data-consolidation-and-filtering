# Agent 执行手册（照着做就能出交付件）

本文件是给**另一个 Hermes agent / 任何自动化执行者**看的操作说明。
`SKILL.md` 讲方法论，这里只讲：**输入长什么样、命令怎么敲、产物怎么验、失败怎么办**。

> 通用原则：不要临场发挥改写脚本，不要猜口径。任何一步的退出码非 0 或校验 FAIL，
> 都按第 9 节的失败矩阵处理；处理不了就停下来向用户要数据或确认口径。

---

## 1. 路线判定（先选一条）

| 用户说法 | 路线 | 主脚本 | 交付物 |
|---|---|---|---|
| 「把这些日报/表格汇总成一张总表」 | **A 宽表** | `long_to_wide.py` | 宽表 csv/xlsx + 明细 + 说明 |
| 「从这些公告/列表里筛出符合条件的」 | **B 筛选** | 按 `SKILL.md` §3 写筛选脚本 | 命中清单 + 增量状态 |
| 「按模板出曲线图」「出四件套」「曲线图+三张表」 | **C 交付** | `xlsx_chart_uniform.py` → `xlsx_charts_to_png.py` → `make_delivery.py` | 5 个文件（见 §6） |

判定不了就问一句：**「你要的是汇总表，还是曲线图+汇总表的交付包？」**
不要默认按 C 跑：C 需要原始 Excel 内嵌图表工作簿，没有它无法开工。

---

## 2. 路径与预检

加载技能后，`skill_view` 会返回 `_source_path`，其目录就是 `SKILL_DIR`。
后面所有命令用绝对路径，**不要依赖当前目录**：

```bash
SKILL_DIR=/opt/data/skills/productivity/data-consolidation-and-filtering   # 换成实际路径
python3 "$SKILL_DIR/scripts/preflight.py" --route C \
    --summary 汇总.xlsx --workbook 原图.xlsx
```

`--summary` 会实际打开工作簿核对表头，`--workbook` 会数 `xl/charts/*.xml`。
两者都可省，但正式交付建议带上——能提前发现「汇总表结构不对」「给的不是内嵌图表工作簿」。

- 退出码 `0`：可以开工（WARN 可接受，例如没中文字体）。
- 退出码 `1`：有 BLOCK。按输出的 `fix` 装依赖后重跑。
- 典型修复：
  ```bash
  python3 -m venv .venv
  .venv/bin/pip install -i https://pypi.tuna.tsinghua.edu.cn/simple openpyxl pillow matplotlib
  .venv/bin/python "$SKILL_DIR/scripts/preflight.py" --route C
  ```
  之后所有脚本都用 `.venv/bin/python` 跑。

依赖分工：`long_to_wide` / `pick_representatives` / `pick_profile_curves` / `xlsx_chart_uniform` 零依赖；
`xlsx_charts_to_png` 需 matplotlib；`make_delivery` 需 openpyxl + Pillow。
**没有 matplotlib 就不能声称出了 PNG/PDF**，只能提出装依赖或改用 Excel 内嵌图。

---

## 3. Route C 输入清单（缺一不可）

开工前逐项确认，缺哪项就直接问用户要，不要用别的数据顶替：

| 输入 | 形态 | 必需 | 说明 |
|---|---|---|---|
| 原始曲线工作簿 | `.xlsx`，内含图表 | 是 | 提供版式基准；一个工作表 = 一组图 |
| 监测数据汇总 | `.xlsx` | 是 | 测项宽表 + 测斜矩阵，结构见 §4 |
| 测斜选点说明 | `.json` | 否 | 不给则每孔按日期等距取 6 期 |
| 项目名 | 字符串 | 是 | 用于全部文件名前缀，例：`110二期X3#井` |
| 输出目录 | 可写路径 | 是 | 建议 `output/<项目>-curves/` |

### 汇总工作簿结构（硬要求）

**测项宽表**（工作表名不要以「测斜」开头）：

```
第 1 行：标题/说明（可为空）
第 2 行：A2 = 点号    B2 = 2024-07-30   C2 = 2024-08-30 …   ← 必须是 YYYY-MM-DD
第 3 行起：A 列 = 点号，其余 = 各期累计值；空就是没测，禁止补 0
```

**测斜矩阵**（工作表名必须以「测斜」开头）：

```
第 1 行：说明（可为空）
第 2 行：A2 = 深度(m)   B2 = 2024-07-30 …
第 3 行起：A 列 = 深度，其余 = 各期累计位移(mm)
工作表名「测斜TST1」→ 孔号取 TST1
```

- 单位按表名自动判断：`倾斜→‰`、`轴力→kN`、其余 `mm`；不确定时用 `--units units.json` 覆盖。
- 表头是合并单元格时，脚本读的是**缓存值**（工作簿由 Excel/WPS 保存过才有值）。

---

## 4. 五步命令（Route C）

以下命令可整段复制，只改 `<...>` 占位符。建议在输出目录的父目录执行。

```bash
SKILL_DIR=/opt/data/skills/productivity/data-consolidation-and-filtering
PY=.venv/bin/python            # 没建 venv 就写 python3
PROJ="110二期X3#井"
WORK=output/110-x3
mkdir -p "$WORK"

# 1) 套版式：先探结构，再按需编辑 chart.json
"$PY" "$SKILL_DIR/scripts/xlsx_chart_uniform.py" --init 原图.xlsx > "$WORK/chart.json"
#    chart.json 不编辑也能跑：--init 默认「每表一组、ref=第 1 张」。
#    只有跨表共用版式时才合并 groups，ref 指向手调好的那张图。
#    图例要显示日期（而不是 45503 这类序列号）：group 里加
#    {"legend_dates": {"row": 2, "from_col": 2, "count": 6}}
"$PY" "$SKILL_DIR/scripts/xlsx_chart_uniform.py" 原图.xlsx "$WORK/统一版式.xlsx" -c "$WORK/chart.json"
test $? -eq 0 || echo "!! 退出码 2 = 有图数据被改动，必须回滚重做"

# 2) 重绘 PNG：auto 会自动区分时程/测斜，测斜落到 png/incl/
"$PY" "$SKILL_DIR/scripts/xlsx_charts_to_png.py" "$WORK/统一版式.xlsx" "$WORK/png" --dpi 220

# 3) （可选）测斜剖面张数自适应，结果写进选点 JSON
"$PY" "$SKILL_DIR/scripts/pick_profile_curves.py" 测斜TST1.csv --format json > "$WORK/TST1.json"

# 4) 四件套（不传 --incl 时自动用 png/incl）
"$PY" "$SKILL_DIR/scripts/make_delivery.py" \
    --summary 汇总.xlsx --groups "$WORK/png/_groups.json" \
    --charts "$WORK/png" --sel "$WORK/sel.json" \
    --title "$PROJ" --out "$WORK/deliverables"

# 5) 验收（这一步不能省）
"$PY" "$SKILL_DIR/scripts/verify_delivery.py" \
    --dir "$WORK/deliverables" --title "$PROJ" --expect-incl 10 --strict
```

关键约定：

- **不要给 `--incl` 传一个空目录/不存在的目录**。旧模板要求文件名含「全过程」；
  新版 `xlsx_charts_to_png.py` 已自动把测斜图放进 `png/incl/` 并补上「全过程」。
- `--time-sheets` 默认 `auto`：轴是「深度」或 orientation=maxMin → 测斜，其余 → 时程。
  自动判错时（例如图名/轴不规范）用 `--time-sheets '图1*,图2*'` 显式指定。
- `xlsx_chart_uniform.py` 退出码 `2` = 检测到 `numCache` 变化，**不要继续下一步**。
- 第 2 步退出码 `2` = 一张时程图都没识别出来；按提示指定 `--time-sheets` 重跑。

### 选点 JSON 格式（`--sel`）

```json
[{"孔号": "TST1", "选定日期": ["2024-07-30", "2024-08-30", "2024-11-08"]}]
```

不给时 `make_delivery` 会对每孔按日期等距取 6 期；`pick_profile_curves.py` 的输出可
直接拼成这个格式（`--format json` 给出 `dates` 字段）。

---

## 5. `_groups.json` 契约（第 2 步产物，第 4 步输入）

```json
[{
  "sheet": "地表沉降",
  "title": "地表沉降 时程曲线 DL1~DL5",
  "file": "地表沉降 时程曲线 DL1~DL5.png",
  "points": ["DL1", "DL2", "DL3", "DL4", "DL5"],
  "dates": ["2024-07-30", "2024-08-30"],
  "series": {"DL1": [0.0, -3.2], "DL2": [0.0, -1.1]}
}]
```

- 一个元素 = 一张图；`points` 必须 ≤5（脚本已按前缀分类、每 5 点一张）。
- `dates` 与 `series[*]` 等长；缺失值写 `null`，不要写 0。
- `file` 是相对 `--charts` 的文件名。
- 不要手写这个文件充数：必须由第 2 步生成，否则图和数据对不上。

---

## 6. 输出契约（验收口径）

目录 `deliverables/` 下必须有且仅有这 5 类文件（另可有日志）：

| 文件 | 内容 | 硬指标 |
|---|---|---|
| `<项目>_数据汇总表.xlsx` | 全部采集数据，测项按自然序 | 行数与原表一致；不补 0 |
| `<项目>_筛选后数据汇总表.xlsx` | 按图分块：上图下表；测斜每孔一块 | 嵌图数 = 时程图数 + 孔数 |
| `<项目>_各点最大累计值汇总表.xlsx` | 每点绝对值最大 + 出现日期；测斜另给深度 | 行数 = 测点数（+孔数） |
| `<项目>_曲线图.pdf` | 封面 + 时程 4 张/页 + 测斜 1 页/孔 | 页数 = 1 + ⌈时程/4⌉ + 孔数 |
| `<项目>_曲线图_PNG.zip` | 全部 PNG + `_索引.csv` | 张数 = 时程 + 孔数 |

图面规则（`verify_delivery` 会查一部分）：图例只写「观测日期」；图名不带 `（mm）`、不带 `1/4`；
曲线一律实线、9 色 + 空心异形标记；测斜横轴落在峰值 1.5~2.5 倍且包住所有点。

---

## 7. 校验与退出码

```bash
"$PY" "$SKILL_DIR/scripts/verify_delivery.py" --dir deliverables --title "$PROJ" \
      --groups png/_groups.json --expect-incl 10
```

| 退出码 | 含义 | 动作 |
|---|---|---|
| `0` | 全部 PASS（WARN 允许） | 可以交付 |
| `2` | 有 FAIL | 按 FAIL 行修复后重跑，**不得带病交付** |
| `1` | 参数错/缺 openpyxl | 修正参数或环境 |

逐项含义：PDF 页数、ZIP 索引张数、筛选表嵌图数、最大表行数、groups 分组与图名。
`--strict` 把 WARN 也当失败，正式对外交付建议加上。

---

## 8. 交付前人工抽检（脚本查不出的）

1. 打开 PDF 第 2 页，确认图内**中文不是方块**、日期不重叠。
2. 抽 2 个测点，把「最大累计值表」的数值与汇总表逐格对一遍。
3. 转灰度看一张 8~9 条曲线的图，标记形状能区分。
4. 图名/页眉里不出现 `mm`、`1/4`、`[数据页…]` 这类残留。
5. 文件名前缀与用户给的项目名完全一致（含井号、二期等字样）。

---

## 9. 失败矩阵

| 症状 | 原因 | 处理 |
|---|---|---|
| `ModuleNotFoundError: matplotlib/PIL/openpyxl` | 依赖没装 | 跑 `preflight.py`，按 fix 建 venv 安装；装完用 venv 的 python |
| `--init` 输出空 groups | 工作簿没有内嵌图表，或图表不在 `xl/charts/` | 确认是「Excel 内嵌图表」工作簿；纯数据表走路线 A |
| `xlsx_chart_uniform` 退出码 2 | 有图数据被改动 | 不要用输出；检查 chart.json 的 ref/分组是否指错 |
| `xlsx_charts_to_png` 退出码 2 | 没识别出时程图 | 用 `--time-sheets '图1*'` 显式指定工作表 |
| `make_delivery` 报「没有识别到测项宽表」 | 表头不在第 2 行，或日期不是 YYYY-MM-DD | 按 §3 调整汇总表；不要改脚本判据 |
| `verify` PDF 页数不符 | 缺孔图或时程图数不对 | 看 `png/incl` 是否为空、`_groups.json` 条数 |
| `verify` 筛选表嵌图数不符 | 测斜文件名不含「全过程」/孔号对不上 | 重新用第 2 步生成；不要手工拷 PNG |
| 图内中文成方块 | 无中文字体 | 放字体到 `~/.fonts/` + `fc-cache -f`，或 `--font 路径` |
| 日期显示 1899/1900 | 图表 x 值不是 Excel 日期序列 | 源图需用日期轴；否则在汇总表按日期列重建图 |
| 测斜图只有 5 条、峰值没体现 | 选点不合理 | 用 `pick_profile_curves.py` 生成 sel.json，别固定 6 条 |
| `PermissionError` 写输出 | 目录只读 | 换 `--out` 到可写目录 |

---

## 10. 完成报告模板（必须按这个说）

```
交付（<项目>）：
- 曲线图 PDF：<N> 页 = 封面 + ⌈<T>/4⌉ + <H> 孔
- 数据汇总表：<X> 个测项 / <Y> 个测点 / <Z> 个测斜孔，<起>~<止>
- 筛选后汇总表：<T+H> 块，嵌图 <T+H> 张
- 最大累计值表：最大 <值><单位>（<点号>，<日期>；测斜：<孔> 深 <d> m）
- PNG 包：<T+H> 张 + 索引
自检：verify_delivery PASS（FAIL 0 / WARN <n>）
异常与处理：<僵尸表/重复实体/缺测/单位存疑…，没有就写“无”>
```

先给结论数字，再给一两句方法；源数据异常单列；附件用 `MEDIA:` 绝对路径发。

---

## 11. 禁止事项

- 不补 0、不插值、不删点来「让曲线好看」。
- 不静默覆盖重复日期；冲突写进 `_conflicts.csv` 并在说明页列出。
- 不用截图/占位图冒充 PNG；没有图就说没图。
- 不改 `verify_delivery` 的判据来「让它通过」。
- 缺数据源、口径冲突、项目名不确定时**停下来问**，不要自己起名或拼数据。
