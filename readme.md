**这是一个自动填表器**  
**This is an auto-fill Excel machine**

它可以根据 Excel 表格中指定单元格（如 E3、E4）的成绩数值，从不同词库中随机选取关键词，自动填入对应的单元格（如 B5、B6、F3、F4），并支持整行数据填充。适用于生成个性化评语、学生反馈等批量处理场景。

## ✨ 功能特点

- 📊 **多学科支持**：一个 Excel 文件可包含多个工作表（如语文、数学、英语），每个工作表独立配置填充规则。
- 🎯 **条件填充**：根据分数所在单元格（例如 E3）的数值，自动选择对应档位的词库（如 ≥75 分使用“优秀”词库，<75 分使用“待改进”词库）。
- 🚫 **无效分数处理**：若分数单元格为空或包含“缺考”“无”等标记，可跳过填充（保持原内容不变）。
- 📥 **批量处理**：递归扫描 `data` 文件夹下所有 `.xlsx` 文件，保持目录结构输出到 `output` 文件夹。
- 🔧 **灵活配置**：所有规则通过 JSON 配置文件定义，无需修改代码。
- 📦 **打包为独立 exe**：提供一键打包脚本（`install.bat` / `install.sh`），可生成无 Python 环境也能运行的 `.exe` 文件。

## 🚀 快速开始

### 1. 准备环境（源码运行）
- Python 3.8 及以上（推荐 3.10+）
- 安装依赖：`pip install openpyxl`

### 2. 准备词库文件
按学科和单元格组织词库，每个词库为 UTF-8 编码的文本文件，每行一个词或短语。  
例如：`wordbanks/英语/B5good.txt`

整行填充词库为 CSV 格式（逗号分隔），例如 `wordbanks/英语/A12-E12.csv`：

### 3. 编写配置文件 `fill_config.json`
```json
{
    "input_dir": "data",
    "output_dir": "output",
    "subjects": [
        {
            "name": "英语",
            "sheet_name": "英语",
            "fill_rules": {
                "B5": {
                    "score_cell": "E3",
                    "thresholds": [75, 50],
                    "wordbanks": ["wordbanks/英语/B5good.txt", "wordbanks/英语/B5med.txt", "wordbanks/英语/B5low.txt"],
                    "skip_on_invalid": true,
                    "invalid_markers": ["缺考", "缺", "无", "请假"]
                },
                "B6": {
                    "score_cell": "E4",
                    "thresholds": [80, 60],
                    "wordbanks": ["wordbanks/英语/B6good.txt", "wordbanks/英语/B6med.txt", "wordbanks/英语/B6low.txt"],
                    "skip_on_invalid": true,
                    "invalid_markers": ["缺考", "缺", "无", "请假"]
                },
                "F3": {
                    "score_cell": "E3",
                    "thresholds": [75, 50],
                    "wordbanks": ["wordbanks/英语/F3good.txt", "wordbanks/英语/F3med.txt", "wordbanks/英语/F3low.txt"],
                    "skip_on_invalid": true,
                    "invalid_markers": ["缺考", "缺", "无", "请假"]
                },
                "F4": {
                    "score_cell": "E4",
                    "thresholds": [80, 60],
                    "wordbanks": ["wordbanks/英语/F4good.txt", "wordbanks/英语/F4med.txt", "wordbanks/英语/F4low.txt"],
                    "skip_on_invalid": true,
                    "invalid_markers": ["缺考", "缺", "无", "请假"]
                },
                "A12:E12": {
                    "file": "wordbanks/英语/A12-E12.csv",
                    "delimiter": ","
                }
            }
        },
        {
            "name": "数学",
            "sheet_name": "数学",
            "fill_rules": {
                "B5": {
                    "score_cell": "E3",
                    "thresholds": [70, 40],
                    "wordbanks": ["wordbanks/数学/B5good.txt", "wordbanks/数学/B5med.txt", "wordbanks/数学/B5low.txt"]
                },
                "A12:E12": {
                    "file": "wordbanks/数学/A12-E12.csv",
                    "delimiter": ","
                }
            }
        }
    ]
}
```

### 4. 运行脚本
```bash
python fill_excel.py
```

## 📦 打包为 EXE（无 Python 环境）

项目提供了 `install.bat` (Windows) 和 `install.sh` (Linux/macOS) 打包脚本。

### Windows
```cmd
install.bat fill_excel.py
```
等待完成，生成的 `fill_excel.exe` 位于 `dist` 文件夹。

### Linux/macOS
```bash
chmod +x install.sh
./install.sh fill_excel.py
```
生成的二进制文件位于 `dist` 目录。

打包后可将整个 `output` 目录和词库文件夹一同分发，无需安装 Python。

## 📂 目录结构示例

```
项目根目录/
├─ fill_excel.py          # 主脚本
├─ fill_config.json       # 配置文件
├─ install.bat            # Windows 打包脚本
├─ install.sh             # Linux/macOS 打包脚本
├─ data/                  # 存放待处理的 Excel 文件
│  ├─ 张三.xlsx
│  └─ 李四.xlsx
├─ output/                # 处理后输出（自动创建）
├─ wordbanks/             # 词库文件夹
│  ├─ 英语/
│  │  ├─ B5good.txt
│  │  ├─ B5med.txt
│  │  ├─ B5low.txt
│  │  ├─ B6good.txt
│  │  ├─ ...
│  │  └─ A12-E12.csv
│  └─ 数学/
│      └─ ...
```

## ⚙️ 配置详解

### 通用设置
- `input_dir`：原始 Excel 文件夹（递归扫描所有 `.xlsx`）
- `output_dir`：输出文件夹（保持相同目录结构）
- `subjects`：学科列表，每个学科包含：
  - `name`：学科名称（仅用于日志）
  - `sheet_name`：Excel 中的工作表名（必须完全一致）
  - `fill_rules`：填充规则字典

### 填充规则
- **单元格填充**（如 `"B5"`）：
  - `score_cell`：读取分数的单元格地址（如 `"E3"`）
  - `thresholds`：阈值数组，从高到低排列，例如 `[75, 50]` 表示：
    - 分数 ≥ 75 → 使用 `wordbanks[0]`
    - 50 ≤ 分数 < 75 → 使用 `wordbanks[1]`
    - 分数 < 50 → 使用 `wordbanks[2]`
  - `wordbanks`：与阈值区间对应的词库文件路径数组（长度为 `len(thresholds)+1`）
  - `skip_on_invalid`：分数无效时是否跳过填充（默认 `true`）
  - `invalid_markers`：识别“无效分数”的关键词列表（如 `["缺考", "无"]`），单元格内容包含任一关键词即视为无效。
- **整行填充**（如 `"A12:E12"`）：
  - `file`：CSV 文件路径，每行代表整行要填入的若干字段
  - `delimiter`：字段分隔符，默认为 `,`

### 分数有效判断逻辑
1. 单元格为空 → 无效
2. 单元格内容包含 `invalid_markers` 中的任一字符串 → 无效
3. 单元格内容无法转换为数字 → 无效
4. 能成功转换为数字 → 有效，用于阈值比较

## 💡 常见问题

### 1. 运行脚本提示 `No module named 'openpyxl'`
- 未安装依赖：执行 `pip install openpyxl`
- 若已打包为 exe，请确保打包时使用了 `--hidden-import=openpyxl` 参数（`install.bat` 已默认添加）。

### 2. Excel 打开 CSV 文件中文乱码
- 将词库文件另存为 **UTF-8 with BOM** 编码（使用记事本或 Notepad++）。
- 或使用 Excel 的“数据 → 从文本/CSV”导入，并选择 UTF-8 编码。

### 3. 如何处理不同学科的独立词库？
- 在 `wordbanks` 路径中使用子文件夹区分，如 `wordbanks/语文/`、`wordbanks/数学/`。
- 配置文件中对每个学科的 `fill_rules` 分别指定对应的词库路径。

### 4. 打包后的 exe 报毒或被拦截
- PyInstaller 打包的文件有时会被杀毒软件误报，添加信任即可。
- 可使用 Nuitka 替代打包，误报率更低。

## 📜 开源许可

本项目基于 **GPL 3.0** 协议开源，详情参见 [LICENSE](LICENSE) 文件。

---

**Enjoy!** 如有问题欢迎提 Issue。
