from docx import Document
from document_parser import DocumentParser
from test_case_generator import TestCaseGenerator
from z3 import *

def main():
    parser = DocumentParser()
    
    # 解析所有文档
    parser.parse_variable_doc("input/Data.docx")
    parser.parse_module_doc("input/Module.docx")
    
    # 获取解析后的数据
    data_manager = parser.data_manager
    
    # 创建测试用例生成器
    test_generator = TestCaseGenerator(data_manager)
    
    # 清空日志文件
    with open('log.txt', 'w', encoding='utf-8') as f:
        f.write("MCDC测试用例生成日志\n")
        f.write("="*50 + "\n\n")
    
    # 处理所有模块
    for module_name, module in data_manager.modules.items():
        test_cases = test_generator.generate_mcdc_cases(module)
        if test_cases:
            test_generator.export_to_excel(test_cases, module)
    
    # 保存所有测试用例
    test_generator.save_workbook("test_cases.xlsx")
if __name__ == "__main__":
    main() 