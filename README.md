# hw-annotator · 作业答案自动批注工具

把「拿到 docx 试卷 → 做题 → 答案红色加粗内联标注到每道题 → 导出 PDF」固化成一条可重复运行的流水线。

两种形态，同一套核心：

- **图形界面版**（`hw_annotate_gui.py`）：打开即窗口，把 `.docx` / `answers.json` / 文件夹拖进去，点「开始转换」即可。
- **命令行版**（`hw_annotate.py`）：适合批量与自动化。

## 标注样式

- 选择题：答案字母写在题干末尾，形如 `（B）`
- 完形填空：答案字母紧跟每个 `____16____` 空格之后
- 阅读表达：答案写在题目下方的横线上
- 全部红色加粗；自动删除文末旧「参考答案」块

## 快速开始（图形界面版）

1. 从 Release 下载 `作业答案标注_v1.0.1_win64.zip` 并解压
2. 双击 `作业答案标注.exe`
3. 拖入试卷（`.docx`）+ 答案（`answers.json`）或整个文件夹
4. 点「开始转换」，在试卷同目录得到 `_含答案.docx` 与 `_含答案.pdf`

## 答案文件格式

```json
{
  "file": "国庆假期作业1.docx",
  "choice":    { "1": "B", "2": "C", "36": "D" },
  "cloze":     { "16": "B", "17": "C" },
  "expression": { "56": "Because her best friend's dad was her role model." }
}
```

编码支持 UTF-8（含/不含 BOM）与 GBK/ANSI。

## 命令行版

```bash
python hw_annotate.py run  试卷.docx -a answers.json [--out 目录]
python hw_annotate.py run  试卷目录   -a answers.json ...      # 批量
python hw_annotate.py validate 试卷.docx -a answers.json      # 只校验不写盘
```

## 技术栈

- Python 3.12
- PyMuPDF（docx→HTML→PDF，含 CJK）
- tkinter + tkinterdnd2（GUI 拖拽）
- PyInstaller（打包）

## 目录结构

```
hw-annotator/
├── hw_annotate.py          # CLI 主程序
├── hw_annotate_gui.py      # GUI 主程序（拖拽界面）
├── hw_annotate.spec        # CLI 打包配置
├── hw_annotate_gui.spec    # GUI 打包配置
├── config.yaml             # 默认配置
├── answers/                # 示例答案
└── samples/                # 示例试卷
```

## 构建

```bash
# 需 Python 3.12 + PyMuPDF + tkinterdnd2 + PyInstaller
python -m PyInstaller hw_annotate.spec          # CLI
python -m PyInstaller hw_annotate_gui.spec      # GUI
```

## 说明

- 系统要求：Windows 10/11（64 位），无需 Word/WPS/LibreOffice，完全离线
- 未做数字签名，首次运行 Windows 可能弹 SmartScreen 提示（「更多信息 → 仍要运行」即可）
