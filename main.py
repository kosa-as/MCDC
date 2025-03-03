from docx import Document
from document_parser import DocumentParser

def main():
    parser = DocumentParser()
    
    # 解析所有文档
    parser.parse_variable_doc("input/Data.docx")
    parser.parse_module_doc("input/Module.docx")
    
    # 获取解析后的数据
    data_manager = parser.data_manager
    
    # 同时输出到终端和文件
    with open('log.txt', 'w', encoding='utf-8') as f:
        # 示例：打印所有变量的详细信息
        # for symbol, variable in data_manager.variables.items():
        #     print(f"\n变量详情:")
        #     print(f"变量名: {variable.name}")
        #     print(f"变量符号: {variable.symbol}")
        #     print(f"类型: {variable.var_type}")
        #     print(f"类型解释: {variable.type_desc}")
        #     print(f"初始值: {variable.initial_value}")
        #     print(f"注释: {variable.comment}")
        #     print(f"标识位: {variable.identifier}")
        #     print(f"取值范围: [{variable.min_value}, {variable.max_value}]")
        
        # 示例：打印所有模块
        for module_name, module in data_manager.modules.items():
            output = f"\n{'='*50}\n"
            output += f"任务名称：{module.name}\n"
            output += f"编号：{module.number}\n"
            output += f"功能：{module.function}\n"
            output += f"前置条件：{module.precondition}\n"
            output += f"输入：{', '.join(module.inputs) if module.inputs else ''}\n"
            output += f"输出：{', '.join(module.outputs) if module.outputs else ''}\n"
            output += f"公式：{module.formula}\n"
            output += f"{'='*50}\n"
            
            # 输出到终端
            print(output)
            # 输出到文件
            f.write(output)

if __name__ == "__main__":
    main() 