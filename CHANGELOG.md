# 更新记录

## 1.3.0 — 2026-11-10

目标：让**其他 Hermes agent 拿到本技能后能独立跑出同样的交付件**，而不是只能读方法论。

### 新增
- `references/agent-runbook.md`：输入契约、五步命令、`_groups.json`/`sel.json` 结构、
  输出契约、验收口径、失败矩阵、完成报告模板、禁止事项。
- `scripts/preflight.py`：开工前自检（Python、依赖、中文字体、脚本完整性、写权限），
  退出码 1 = 有阻塞，并给出可照做的 fix 命令。
- `scripts/verify_delivery.py`：交付验收（5 个文件、PDF 页数、ZIP 张数、筛选表嵌图数、
  最大表行数、分组 ≤5 点、图名残留、日期序列异常），退出码 2 = 不得交付。
- `tests/test_delivery_smoke.py`：端到端冒烟（生成最小输入 → 四件套 → 验收）。
- `SKILL.md` §0 执行协议：路线判定、路径约定、硬约束。

### 修复（都会导致其他 agent 静默产出错件）
- **测斜 PNG 不再漏进交付件**：`xlsx_charts_to_png.py` 现在把测斜图写入 `<outdir>/incl/`
  并自动补「全过程」；`make_delivery.py` 的 `--incl` 改为可选，缺省自动取 `<charts>/incl`，
  目录不存在/无匹配时明确 WARN，不再静默 0 张。
- **孔图匹配不再依赖旧文件名模板**：按孔号匹配，优先「全过程」，避免“筛选后汇总表没有测斜图”。
- **自动区分时程/测斜**：`--time-sheets` 默认 `auto`，按「深度轴 / orientation=maxMin」判别，
  不再默认 `图1*` 把所有非图1工作表都当成测斜。
- **中文字体自动发现**：扫描 `~/.fonts`、Noto CJK、文泉驿等，支持 `--font`，减少豆腐块。
- `make_delivery.py` 汇总表结构不符时给出明确错误，而不是产出空表。
- 文档中的一条龙命令统一改为 `--incl out/png/incl`（原 `out/incl` 会让测斜曲线全部丢失）。

### 文档
- `SKILL.md` 命令全部改为 `$SKILL_DIR` 绝对路径，补预检/验收步骤与硬约束。
- `README.md` 更新依赖、路线表、安装路径与验收流程。
- `references/delivery-workflow.md` 同步新的自动分类与目录约定。

## 1.2.0

- 交付节奏：日报打包发送后直接接着跑数据处理。
- 新增 `pick_profile_curves.py` 自适应剖面张数、`make_delivery.py` 四件套。

## 1.1.0

- 新增 Excel 内嵌图表批量套版式与 PNG 重绘。

## 1.0.0

- 长表→宽表、代表点抽样、筛选与增量去重方法论。
