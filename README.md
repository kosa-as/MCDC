# MCDC测试用例生成工具

这是一个基于MCDC（Modified Condition/Decision Coverage）准则的测试用例自动生成工具。该工具可以解析Word文档中的变量和模块定义，并生成满足MCDC覆盖准则的测试用例。

## 环境要求

- Python 3.8或更高版本
- 依赖包：
  - python-docx
  - z3-solver
  - openpyxl

## 安装

1. 克隆本仓库：
```bash
git clone https://github.com/kosa-as/MCDC.git
cd MCDC
```

2. 安装依赖：
```bash
pip install python-docx z3-solver openpyxl
```

## 使用方法

1. 准备输入文件：
   - 在`input`目录下放置`Data.docx`（变量定义文档）
   - 在`input`目录下放置`Module.docx`（模块定义文档）

2. 运行主程序：
```bash
python main.py
```

3. 查看结果：
   - 生成的测试用例将保存在`test_cases.xlsx`文件中
   - 程序运行日志将保存在`log.txt`文件中

## 文件结构

- `main.py`：主程序入口
- `document_parser.py`：文档解析器
- `test_case_generator.py`：测试用例生成器
- `data_structures.py`：数据结构定义
- `input/`：输入文档目录
  - `Data.docx`：变量定义文档
  - `Module.docx`：模块定义文档
