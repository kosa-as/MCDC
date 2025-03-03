from z3 import *
import re
from openpyxl import Workbook
import ast
import sys
from contextlib import contextmanager

@contextmanager
def log_to_file(log_file):
    """将输出同时写入到文件和终端的上下文管理器"""
    class TeeWriter:
        def __init__(self, *files):
            self.files = files
        
        def write(self, obj):
            for f in self.files:
                f.write(obj)
                f.flush()  # 立即刷新缓冲区
        
        def flush(self):
            for f in self.files:
                f.flush()
    
    # 保存原始的标准输出
    original_stdout = sys.stdout
    
    try:
        # 创建同时写入到终端和文件的writer
        tee = TeeWriter(sys.stdout, log_file)
        sys.stdout = tee
        yield
    finally:
        # 恢复原始的标准输出
        sys.stdout = original_stdout

class TestCaseGenerator:
    def __init__(self, data_manager):
        self.data_manager = data_manager
        self.workbook = Workbook()
        self.sheet = self.workbook.active
        self._init_sheet()
        
    def _init_sheet(self):
        """初始化Excel表头"""
        headers = ["需求编号", "模块名称", "前置条件", "判断条件", "测试用例", "预期结果"]
        for col, header in enumerate(headers, 1):
            self.sheet.cell(row=1, column=col, value=header)
            
    def _parse_condition(self, condition):
        """解析条件表达式，返回变量列表和处理后的条件"""
        # 先处理特殊函数调用
        def handle_last_function(match):
            var_name = match.group(1)  # 获取括号中的变量名
            return f"_{var_name}_"  # 返回新的变量名格式
        
        # 替换 last(X) 为 _X_
        condition = re.sub(r'last\((\w+)\)', handle_last_function, condition)
        
        # 先移除所有空格，以便更好地处理操作符
        condition = ''.join(condition.split())
        
        # 替换C语言操作符为Python操作符，注意顺序很重要
        # 先处理多字符操作符，再处理单字符操作符
        condition = re.sub(r'<=', ' <= ', condition)  # 先处理 <=
        condition = re.sub(r'>=', ' >= ', condition)  # 再处理 >=
        condition = re.sub(r'==', ' == ', condition)
        condition = re.sub(r'!=', ' != ', condition)
        condition = re.sub(r'&&', ' and ', condition)
        condition = re.sub(r'\|\|', ' or ', condition)
        condition = re.sub(r'(?<![=<>!])=(?![=])', ' == ', condition)  # 单独的 = 转换为 ==
        condition = re.sub(r'(?<![<>])>(?![=])', ' > ', condition)  # 单独的 >
        condition = re.sub(r'(?<![<>])<(?![=])', ' < ', condition)  # 单独的 <
        condition = re.sub(r'!(?![=])', ' not ', condition)  # 单独的 !
        
        # 规范化空格
        condition = ' '.join(condition.split())
        
        # 提取所有标识符
        identifiers = set(re.findall(r'[a-zA-Z_][a-zA-Z0-9_]*', condition))
        # 移除Python关键字
        identifiers = identifiers - {'and', 'or', 'not', 'True', 'False'}
        
        # 分离变量和常量，并获取它们的范围信息
        variables = []
        processed_condition = condition
        
        for identifier in identifiers:
            # 检查是否是last变量（以_开头和结尾）
            is_last_var = identifier.startswith('_') and identifier.endswith('_')
            original_var_name = identifier[1:-1] if is_last_var else identifier
            
            # 检查是否是常量
            if original_var_name in self.data_manager.constants:
                # 将常量也作为变量处理，但标记为常量
                variables.append({
                    'name': identifier,
                    'min_value': float(self.data_manager.constants[original_var_name]),
                    'max_value': float(self.data_manager.constants[original_var_name]),
                    'var_type': 'constant',
                    'is_last': is_last_var,
                    'original_var': original_var_name,
                    'constant_value': float(self.data_manager.constants[original_var_name])
                })
            else:
                # 检查是否是已定义的变量
                var_found = False
                for var_symbol, var_obj in self.data_manager.variables.items():
                    if original_var_name == var_symbol:
                        # 如果是last变量，使用新的变量名但保持相同的范围和类型
                        var_name = identifier if is_last_var else original_var_name
                        variables.append({
                            'name': var_name,
                            'min_value': var_obj.min_value,
                            'max_value': var_obj.max_value,
                            'var_type': var_obj.var_type,
                            'is_last': is_last_var,
                            'original_var': original_var_name
                        })
                        var_found = True
                        break
                if not var_found:
                    print(f"警告：未找到变量或常量定义：{original_var_name}")
        
        print(f"解析后的Python条件表达式: {processed_condition}")  # 调试输出
        print(f"变量及其范围: {variables}")  # 调试输出
        return variables, processed_condition
        
    def _create_z3_vars(self, variables):
        """为每个变量创建Z3 Bool变量"""
        return {var: Bool(var) for var in variables}
        
    def _generate_mcdc_conditions(self, condition, variables):
        """生成MCDC测试条件"""
        # 创建一个基本的解析器来处理特殊函数和复杂表达式
        def preprocess_condition(cond):
            """预处理条件，处理特殊函数和复杂表达式"""
            # 替换&&和||为and和or
            cond = cond.replace('&&', ' and ').replace('||', ' or ')
            
            # 处理abs函数
            def replace_abs(match):
                expr = match.group(1)
                return f"(({expr}) if ({expr}) >= 0 else -({expr}))"
            cond = re.sub(r'abs\(([^)]+)\)', replace_abs, cond)
            
            # 处理duration函数 - 我们将其视为常量True，因为它是时序相关的
            def replace_duration(match):
                return "True"
            cond = re.sub(r'duration\([^)]*\)', replace_duration, cond)
            
            # 移除函数中不必要的参数（如ms）
            cond = re.sub(r',\s*ms\s*,', ',', cond)
            
            return cond
        
        # 预处理条件
        try:
            processed_condition = preprocess_condition(condition)
            print(f"预处理后的条件: {processed_condition}")
        except Exception as e:
            print(f"预处理条件时出错: {str(e)}")
            return []
        
        # 创建Z3变量
        z3_vars = {}
        for var in variables:
            var_name = var['name']
            if var.get('var_type') == 'constant':
                # 对于常量，直接使用其值
                z3_vars[var_name] = IntVal(int(float(var['constant_value'])))
            else:
                # 对于变量，使用Variable中定义的类型
                original_var = self.data_manager.variables[var['original_var']]
                if original_var.var_type == 'int':
                    z3_vars[var_name] = Int(var_name)
                else:
                    z3_vars[var_name] = Real(var_name)
        
        # 创建求解器
        s = Solver()
        
        # 为每个变量添加范围约束
        for var in variables:
            var_name = var['name']
            if not var.get('var_type') == 'constant':
                var_expr = z3_vars[var_name]
                original_var = self.data_manager.variables[var['original_var']]
                if original_var.var_type == 'int':
                    s.add(var_expr >= int(original_var.min_value))
                    s.add(var_expr <= int(original_var.max_value))
                else:
                    s.add(var_expr >= float(original_var.min_value))
                    s.add(var_expr <= float(original_var.max_value))
        
        # 直接使用Z3解析条件
        def parse_z3_condition(condition):
            # 将条件转换为可解析的表达式
            try:
                # 替换所有变量名为z3变量
                for var_name, var_expr in z3_vars.items():
                    # 确保我们匹配完整的变量名（不匹配子字符串）
                    pattern = r'\b' + var_name + r'\b'
                    condition = re.sub(pattern, f'z3_vars["{var_name}"]', condition)
                
                # 替换逻辑操作符
                condition = condition.replace(' and ', ' & ').replace(' or ', ' | ').replace(' not ', ' ~ ')
                
                # 替换比较操作符
                condition = condition.replace('!=', '!=').replace('==', '==').replace('>=', '>=').replace('<=', '<=')
                
                # 添加Z3函数
                scope = {
                    'z3_vars': z3_vars,
                    'And': And,
                    'Or': Or,
                    'Not': Not,
                    'If': If,
                    'True': BoolVal(True),
                    'False': BoolVal(False)
                }
                
                # 调试输出
                print(f"解析的Z3条件: {condition}")
                
                # 安全地评估表达式
                result = eval(condition, scope)
                print(f"成功解析为Z3表达式: {result}")
                return result
            except Exception as e:
                print(f"Z3条件解析错误: {str(e)}")
                # 提供更详细的错误信息，显示当前作用域中的变量
                print(f"变量作用域: {list(z3_vars.keys())}")
                raise
        
        try:
            # 尝试解析条件
            base_expr = parse_z3_condition(processed_condition)
        except Exception as e:
            print(f"无法解析条件: {processed_condition}")
            print(f"错误: {str(e)}")
            return []
        
        # 生成测试条件
        test_conditions = []
        for var in variables:
            var_name = var['name']
            # 跳过常量，因为它们不能作为MCDC条件的变量
            if var.get('var_type') == 'constant':
                continue
            
            # 获取变量的Z3表达式
            var_expr = z3_vars[var_name]
            
            # 使用Z3求解器测试每个变量是否会影响条件
            # 首先尝试找到一个使条件为True的变量值
            s.push()
            s.add(base_expr)
            if s.check() == sat:
                model_true = s.model()
                s.pop()
                
                # 然后尝试找到相同条件下，改变当前变量使条件为False的值
                s.push()
                var_value_true = model_true.eval(var_expr)
                # 添加约束：所有其他变量保持相同值
                for other_var in variables:
                    other_name = other_var['name']
                    if other_name != var_name and not other_var.get('var_type') == 'constant':
                        other_expr = z3_vars[other_name]
                        if other_name in [str(d) for d in model_true.decls()]:
                            s.add(other_expr == model_true.eval(other_expr))
                
                # 添加当前变量取不同值的约束
                s.add(var_expr != var_value_true)
                s.add(Not(base_expr))
                
                if s.check() == sat:
                    model_false = s.model()
                    
                    # 创建测试用例
                    test_case_true = {}
                    test_case_false = {}
                    
                    # 填充测试用例
                    for test_var in variables:
                        test_name = test_var['name']
                        if test_var.get('var_type') == 'constant':
                            # 常量使用固定值
                            value = float(test_var['constant_value'])
                            test_case_true[test_name] = value
                            test_case_false[test_name] = value
                        else:
                            # 变量使用模型中的值
                            test_expr = z3_vars[test_name]
                            if test_name == var_name:
                                # 当前变量使用不同的值
                                true_val = model_true.eval(test_expr)
                                false_val = model_false.eval(test_expr)
                                
                                # 确保我们得到具体值而不是表达式
                                if is_int(true_val):
                                    test_case_true[test_name] = true_val.as_long()
                                else:
                                    test_case_true[test_name] = float(true_val.as_decimal(10))
                                    
                                if is_int(false_val):
                                    test_case_false[test_name] = false_val.as_long()
                                else:
                                    test_case_false[test_name] = float(false_val.as_decimal(10))
                            else:
                                # 其他变量使用相同的值（来自model_true）
                                if test_name in [str(d) for d in model_true.decls()]:
                                    val = model_true.eval(test_expr)
                                    if is_int(val):
                                        value = val.as_long()
                                    else:
                                        value = float(val.as_decimal(10))
                                    test_case_true[test_name] = value
                                    test_case_false[test_name] = value
                                else:
                                    # 如果模型中没有值，使用变量范围的中点
                                    mid = (float(test_var['min_value']) + float(test_var['max_value'])) / 2
                                    if self.data_manager.variables[test_var['original_var']].var_type == 'int':
                                        mid = int(mid)
                                    test_case_true[test_name] = mid
                                    test_case_false[test_name] = mid
                    
                    # 添加测试条件（一对使条件结果不同的测试用例）
                    test_conditions.append((test_case_true, test_case_false))
                
                s.pop()
            else:
                s.pop()
        
        return test_conditions
        
    def generate_mcdc_cases(self, module):
        """为模块生成MCDC测试用例"""
        with open('log.txt', 'a', encoding='utf-8') as log_file:
            with log_to_file(log_file):
                print(f"\n{'='*50}")
                print(f"处理模块: {module.name}")
                
                if not module.formula or 'if' not in module.formula:
                    print("模块中没有if语句，跳过")
                    return None
                
                def find_matching_parenthesis(text, start_pos):
                    """找到匹配的括号"""
                    count = 1
                    pos = start_pos
                    while count > 0 and pos < len(text):
                        pos += 1
                        if pos >= len(text):
                            break
                        if text[pos] == '(':
                            count += 1
                        elif text[pos] == ')':
                            count -= 1
                    return pos if count == 0 else -1
                
                def extract_if_conditions(text):
                    """提取if条件，正确处理嵌套括号"""
                    conditions = []
                    i = 0
                    while i < len(text):
                        if text[i:i+2] == 'if':
                            # 找到if后的第一个左括号
                            left_paren = text.find('(', i)
                            if left_paren != -1:
                                # 计算嵌套括号
                                count = 1
                                right_paren = left_paren + 1
                                while count > 0 and right_paren < len(text):
                                    if text[right_paren] == '(':
                                        count += 1
                                    elif text[right_paren] == ')':
                                        count -= 1
                                    right_paren += 1
                                
                                if count == 0:
                                    # 提取完整的条件，包括所有嵌套的括号
                                    condition = text[left_paren:right_paren].strip()
                                    # 移除最外层的括号
                                    if condition.startswith('(') and condition.endswith(')'):
                                        condition = condition[1:-1].strip()
                                    conditions.append(condition)
                                    i = right_paren
                        i += 1
                    return conditions
                
                # 使用新的方法提取if条件
                if_conditions = extract_if_conditions(module.formula)
                print(f"提取到的if条件: {if_conditions}")
                
                if not if_conditions:
                    print("没有找到有效的if条件")
                    return None
                
                test_cases = []
                for idx, condition in enumerate(if_conditions):
                    print(f"\n处理条件 {idx + 1}: {condition}")
                    variables, processed_condition = self._parse_condition(condition)
                    print(f"解析后的条件: {processed_condition}")
                    print(f"识别到的变量: {variables}")
                    
                    # 过滤变量：优先使用输入变量，对于非输入变量检查是否在变量定义中
                    input_vars = []
                    for var in variables:
                        var_name_to_check = var['original_var'] if var.get('is_last') else var['name']
                        if var_name_to_check in module.inputs:
                            # 如果是输入变量，直接添加
                            input_vars.append(var)
                        elif var_name_to_check in self.data_manager.variables:
                            # 如果是已定义的变量但不是输入变量，也添加
                            input_vars.append(var)
                        elif var_name_to_check in self.data_manager.constants:
                            # 如果是常量，也添加
                            input_vars.append(var)
                    
                    print(f"处理的变量: {[var['name'] for var in input_vars]}")
                    
                    if not input_vars:
                        print("没有找到有效变量，跳过")
                        continue
                    
                    # 生成MCDC测试用例
                    mcdc_conditions = self._generate_mcdc_conditions(processed_condition, input_vars)
                    print(f"生成的MCDC条件数量: {len(mcdc_conditions)}")
                    
                    for test_case, opposite_case in mcdc_conditions:
                        # 创建完整的评估环境，包括测试用例值和常量
                        eval_env = {**test_case, **self.data_manager.constants}
                        opposite_eval_env = {**opposite_case, **self.data_manager.constants}
                        
                        test_cases.append({
                            "编号": module.number,
                            "模块名称": module.name,
                            "前置条件": module.precondition,
                            "判断条件": condition,  # 使用原始条件
                            "测试用例": self._format_test_case(test_case),
                            "预期结果": "True" if eval(processed_condition, {"__builtins__": {}}, eval_env) else "False"
                        })
                        
                        test_cases.append({
                            "编号": module.number,
                            "模块名称": module.name,
                            "前置条件": module.precondition,
                            "判断条件": condition,  # 使用原始条件
                            "测试用例": self._format_test_case(opposite_case),
                            "预期结果": "True" if eval(processed_condition, {"__builtins__": {}}, opposite_eval_env) else "False"
                        })
                
                return test_cases
    
    def _format_condition(self, original_condition, test_case):
        """格式化判断条件，显示具体的取反情况"""
        # 直接返回原始条件，不做任何替换
        return original_condition.strip()
    
    def _format_test_case(self, test_case):
        """格式化测试用例为字符串"""
        parts = []
        
        # 添加变量值
        for var, val in sorted(test_case.items()):  # 排序以保持输出顺序一致
            # 如果是last变量，显示原始变量名
            if var.startswith('_') and var.endswith('_'):
                display_name = f"last({var[1:-1]})"  # 转换回 last(X) 格式
            else:
                display_name = var
            parts.append(f"{display_name}={val}")
        
        return ", ".join(parts)
    
    def export_to_excel(self, test_cases, module):
        """将测试用例导出到Excel"""
        with open('log.txt', 'a', encoding='utf-8') as log_file:
            with log_to_file(log_file):
                print(f"\n生成的测试用例:")
                for test_case in test_cases:
                    # 打印到日志
                    print(f"- 条件: {test_case['判断条件']}")
                    print(f"  测试用例: {test_case['测试用例']}")
                    print(f"  预期结果: {test_case['预期结果']}")
                    
                    # 添加到Excel
                    row = [
                        test_case["编号"],
                        test_case["模块名称"],
                        test_case["前置条件"],
                        test_case["判断条件"],
                        test_case["测试用例"],
                        test_case["预期结果"]
                    ]
                    self.sheet.append(row)
    
    def save_workbook(self, filename):
        """保存Excel工作簿"""
        self.workbook.save(filename) 