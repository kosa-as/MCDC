# 测试用例处理工具

这是一个用于处理测试用例的工具集，可以从Excel文件中提取测试用例，并生成格式化的输出结果。

## 文件说明

- `main.py`: 主程序，按顺序执行所有处理步骤
- `testcase_parser.py`: 解析Excel测试用例文件，提取模块信息和常量
- `generate_testcase.py`: 生成最终的测试用例Excel文件

## 环境要求

- Python 3.8
- 依赖库:
  - pandas
  - openpyxl
  - python-docx
  - z3-solver
  - re

## 安装依赖

```bash
pip install pandas openpyxl python-docx z3-solver
```

## 使用方法

1. 将输入文件放在`input`目录下:
   - `testcase.xlsx`: 测试用例Excel文件
   - `module.docx`: 模块信息Word文档
   - `data.docx`: 数据常量Word文档

2. 运行主程序:
   ```bash
   python main.py
   ```

3. 输出文件将生成在`output`目录下:
   - `generated_testcases.xlsx`: 最终生成的测试用例Excel文件
   - 其他中间JSON文件

## 处理流程

1. 解析Excel测试用例文件，生成JSON格式数据
2. 合并相同需求编号和模块名称的测试用例
3. 解析模块文档，提取模块信息
4. 匹配测试用例和模块信息，提取结果
5. 解析数据文档，提取常量
6. 生成最终的测试用例Excel文件 