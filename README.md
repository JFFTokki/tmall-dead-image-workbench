# 死图判断工作台

一个用于批量检测 Excel 商品图片中“死图/异常图”的本地 Windows 工作台。项目把掌柜软件导出、图片规则判断、人工复核学习、结果拆分交付整合到一个桌面应用里，尽量减少手工跑命令和反复整理结果表的成本。

## 功能概览

- 图形化工作台：通过 `启动死图判断工作台.bat` 打开本地桌面界面。
- Excel 图片检测：读取商品 Excel 中嵌入的主图/长图列，识别疑似死图底部状态条。
- 掌柜软件导出：可按商品 ID 批量调用掌柜软件导出图片元数据 Excel。
- 断点与日志：长任务会保留 checkpoint、运行日志和过程结果，方便排查与续看。
- 结果交付：输出主结果 Excel，并可拆出问题图片明细表给人工复核。
- 人工反馈学习：可读取人工批注结果，把误判样本加入本地反馈模型。
- 可选 LLM 复核：支持通过 OpenAI 兼容接口启用视觉模型二次判断。

## 项目结构

```text
.
├── 启动死图判断工作台.bat        # Windows 一键启动入口
├── dead_image_workbench_modern.py # CustomTkinter 桌面工作台
├── dead_image_workbench_core.py   # 工作台共享逻辑：路径、进度解析、结果拆分
├── maijsoft_export.py             # 掌柜软件按商品 ID 批量导出脚本
├── app_brainstorm_preview.html    # 早期 UI 流程草图
└── 死图判断源数据/
    └── detect_dead_images.py      # 核心死图检测脚本
```


## 环境要求

- Windows 10/11
- Python 3.10+
- 依赖库：
  - `customtkinter`
  - `openpyxl`
  - `pandas`
  - `numpy`
  - `Pillow`

安装依赖示例：

```bash
py -m pip install customtkinter openpyxl pandas numpy pillow
```

如果你本地有 `vendor/python/` 目录，工作台会自动优先加载其中的依赖；该目录默认不提交到 git。

## 快速开始

### 方式一：桌面工作台

双击运行：

```text
启动死图判断工作台.bat
```

或在项目根目录执行：

```bash
py dead_image_workbench_modern.py
```

工作台里可以选择：

1. 已有 Excel 文件；或商品 ID 文件 + HAR/Cookie 导出源数据。
2. 输出目录。
3. 示例/批注 Excel（可选，用于版式和人工反馈）。
4. 是否启用 LLM 二次判断。
5. 开始运行并查看日志、进度和结果文件。

### 方式二：直接检测 Excel

```bash
py "死图判断源数据/detect_dead_images.py" "输入文件.xlsx" --output-dir dead_image_output
```

常用参数：

```bash
py "死图判断源数据/detect_dead_images.py" "输入文件.xlsx" \
  --output-dir dead_image_output \
  --output-xlsx "输出结果.xlsx" \
  --id-column 商品ID \
  --image-columns "主图1,主图2,3:4长图1,3:4长图2" \
  --batch-size 500
```

可选人工反馈参数：

```bash
--template-xlsx "示例输出.xlsx"
--correction-xlsx "人工复核批注.xlsx"
--feedback-model "dead_image_output/models/current/dead_image_model_feedback.json"
```

可选 LLM 参数：

```bash
--llm-enabled \
--llm-api-url "https://你的-openai兼容接口/v1/chat/completions" \
--llm-api-key "你的API Key" \
--llm-model "你的视觉模型名"
```

也可以用环境变量传 API Key，避免出现在命令历史里：

```bash
set DEAD_IMAGE_LLM_API_KEY=你的API Key
```

## 掌柜软件导出

`maijsoft_export.py` 用于按商品 ID 批量导出图片元数据 Excel。它需要你提供已登录浏览器的 Cookie，或从已登录会话导出的 HAR 文件。

```bash
py maijsoft_export.py --har tb.maijsoft.cn.har --ids-file ids.txt --output-dir maijsoft_exports
```

也可以直接传 Cookie：

```bash
py maijsoft_export.py --cookie-file cookie.txt --ids-file ids.txt --output-dir maijsoft_exports
```

常用参数：

- `--ids-file`：商品 ID 文本文件，每行一个，或用空格/逗号分隔。
- `--id`：直接传单个商品 ID，可重复传入。
- `--batch-size`：每批商品 ID 数量，最大 10000。
- `--include-extra-fields`：额外导出标题、价格、库存、商家编码。
- `--dry-run`：只预览分批和字段，不实际提交导出。

示例：

```bash
py maijsoft_export.py --ids-file ids.txt --batch-size 5000 --dry-run
```

## 输入与输出

### 输入 Excel

检测脚本默认按以下逻辑读取：

- 商品 ID 列：默认 `商品ID`，可用 `--id-column` 修改。
- 图片列：默认读取脚本内置的常用图片列名，可用 `--image-columns` 指定。
- 工作表：默认读取第一个/当前工作表，可用 `--sheet-name` 指定。

### 输出文件

输出目录通常在 `dead_image_output/` 或工作台选择的目录下，包括：

- 主结果 Excel：包含每个商品/图片列的检测结果。
- 问题图片 Excel：便于人工集中复核疑似死图。
- checkpoint JSONL：阶段性检测明细。
- 运行日志：用于排查长任务中断、接口异常、图片解析问题。

## AI 辅助说明

本项目在开发过程中使用了 Claude Code 和 OpenAI Codex 提供的辅助。
