from z3 import *
import re
from openpyxl import Workbook
import ast
import sys
from contextlib import contextmanager
import traceback

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
        # 预处理条件
        def preprocess_condition(cond):
            """预处理条件，处理特殊函数和复杂表达式"""
            # 替换&&和||为and和or
            cond = cond.replace('&&', ' and ').replace('||', ' or ')
            
            # 替换特殊字符，如破折号为减号
            cond = cond.replace('–', '-')  # 替换破折号为减号
            
            # 处理abs函数
            def replace_abs(match):
                expr = match.group(1)
                return f"(({expr}) if ({expr}) >= 0 else -({expr}))"
            cond = re.sub(r'abs\(([^)]+)\)', replace_abs, cond)
            
            # 处理duration函数 - 将其视为常量True
            def replace_duration(match):
                return "True"
            cond = re.sub(r'duration\([^)]*\)', replace_duration, cond)
            
            # 移除函数中不必要的参数
            cond = re.sub(r',\s*ms\s*,', ',', cond)
            
            return cond
        
        # 预处理条件
        processed_condition = preprocess_condition(condition)
        print(f"预处理后的条件: {processed_condition}")
        
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
                elif original_var.var_type == 'bool':
                    z3_vars[var_name] = Bool(var_name)
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
                
                if original_var.var_type == 'bool':
                    # 对布尔变量不添加数值范围约束
                    pass
                elif original_var.var_type == 'int':
                    s.add(var_expr >= int(original_var.min_value))
                    s.add(var_expr <= int(original_var.max_value))
                else:
                    s.add(var_expr >= float(original_var.min_value))
                    s.add(var_expr <= float(original_var.max_value))
        
        # 提取原子条件和操作符信息
        def extract_expression_structure(expr_str):
            """提取表达式结构，包括原子条件和连接符"""
            # 替换所有表达式为占位符，以便后续解析
            placeholders = {}
            atomic_conditions = []
            
            # 递归解析表达式
            def parse_recursive(expr, level=0):
                # 去除表达式两端的括号
                expr = expr.strip()
                if expr.startswith('(') and expr.endswith(')') and is_balanced_parentheses(expr[1:-1]):
                    return parse_recursive(expr[1:-1], level)
                
                # 查找顶层的连接符
                top_and_pos = find_top_level_operator(expr, ' and ')
                top_or_pos = find_top_level_operator(expr, ' or ')
                
                if top_or_pos:  # 优先处理OR，因为它的优先级较低
                    # 以OR分割表达式
                    parts = []
                    last_pos = 0
                    for pos in top_or_pos:
                        parts.append(expr[last_pos:pos])
                        last_pos = pos + 4  # ' or '的长度
                    parts.append(expr[last_pos:])
                    
                    # 递归处理每个部分
                    sub_exprs = [parse_recursive(part, level+1) for part in parts]
                    
                    # 组合成OR表达式
                    placeholder = f"OR_EXPR_{level}"
                    placeholders[placeholder] = ('or', sub_exprs)
                    return placeholder
                    
                elif top_and_pos:  # 然后处理AND
                    # 以AND分割表达式
                    parts = []
                    last_pos = 0
                    for pos in top_and_pos:
                        parts.append(expr[last_pos:pos])
                        last_pos = pos + 5  # ' and '的长度
                    parts.append(expr[last_pos:])
                    
                    # 递归处理每个部分
                    sub_exprs = [parse_recursive(part, level+1) for part in parts]
                    
                    # 组合成AND表达式
                    placeholder = f"AND_EXPR_{level}"
                    placeholders[placeholder] = ('and', sub_exprs)
                    return placeholder
                    
                else:  # 原子条件
                    # 解析比较表达式
                    if '==' in expr:
                        left, right = [s.strip() for s in expr.split('==', 1)]
                        atomic_conditions.append((left, '==', right))
                    elif '!=' in expr:
                        left, right = [s.strip() for s in expr.split('!=', 1)]
                        atomic_conditions.append((left, '!=', right))
                    elif '>=' in expr:
                        left, right = [s.strip() for s in expr.split('>=', 1)]
                        atomic_conditions.append((left, '>=', right))
                    elif '<=' in expr:
                        left, right = [s.strip() for s in expr.split('<=', 1)]
                        atomic_conditions.append((left, '<=', right))
                    elif '>' in expr:
                        left, right = [s.strip() for s in expr.split('>', 1)]
                        atomic_conditions.append((left, '>', right))
                    elif '<' in expr:
                        left, right = [s.strip() for s in expr.split('<', 1)]
                        atomic_conditions.append((left, '<', right))
                    
                    # 返回原子条件的索引
                    placeholder = f"ATOM_{len(atomic_conditions)-1}"
                    return placeholder
            
            # 查找顶层操作符位置
            def find_top_level_operator(expr, op):
                positions = []
                level = 0
                for i in range(len(expr) - len(op) + 1):
                    if expr[i] == '(':
                        level += 1
                    elif expr[i] == ')':
                        level -= 1
                    elif level == 0 and expr[i:i+len(op)] == op:
                        positions.append(i)
                return positions
            
            # 检查括号平衡
            def is_balanced_parentheses(s):
                count = 0
                for char in s:
                    if char == '(':
                        count += 1
                    elif char == ')':
                        count -= 1
                        if count < 0:
                            return False
                return count == 0
            
            # 解析表达式结构
            root = parse_recursive(expr_str)
            
            return atomic_conditions, placeholders, root
        
        # 提取表达式结构
        atomic_conditions, expr_structure, root_expr = extract_expression_structure(processed_condition)
        print(f"提取的原子条件: {atomic_conditions}")
        print(f"表达式结构: {expr_structure}")
        print(f"根表达式: {root_expr}")
        
        # 构建Z3表达式
        def build_z3_expression(expr_id, structure, atomic_conds):
            if expr_id.startswith('ATOM_'):
                idx = int(expr_id.split('_')[1])
                left, op, right = atomic_conds[idx]
                
                # 处理复合表达式如H-H_TO
                def resolve_variable(var_name):
                    if var_name in z3_vars:
                        return z3_vars[var_name]
                    
                    # 尝试解析复合表达式如H-H_TO
                    for op_char in ['+', '-', '*', '/']:
                        if op_char in var_name:
                            parts = var_name.split(op_char)
                            if all(part.strip() in z3_vars for part in parts):
                                if op_char == '+':
                                    return z3_vars[parts[0].strip()] + z3_vars[parts[1].strip()]
                                elif op_char == '-':
                                    return z3_vars[parts[0].strip()] - z3_vars[parts[1].strip()]
                                elif op_char == '*':
                                    return z3_vars[parts[0].strip()] * z3_vars[parts[1].strip()]
                                elif op_char == '/':
                                    return z3_vars[parts[0].strip()] / z3_vars[parts[1].strip()]
                
                    print(f"警告: 无法解析变量 {var_name}")
                    return None
                
                left_expr = resolve_variable(left)
                right_expr = resolve_variable(right)
                
                if left_expr is None or right_expr is None:
                    print(f"警告: 无法解析表达式 {left} {op} {right}")
                    return BoolVal(True)
                
                if op == '==':
                    return left_expr == right_expr
                elif op == '!=':
                    return left_expr != right_expr
                elif op == '>=':
                    return left_expr >= right_expr
                elif op == '<=':
                    return left_expr <= right_expr
                elif op == '>':
                    return left_expr > right_expr
                elif op == '<':
                    return left_expr < right_expr
            else:
                op_type, sub_exprs = structure[expr_id]
                sub_results = [build_z3_expression(sub_expr, structure, atomic_conds) for sub_expr in sub_exprs]
                
                if op_type == 'and':
                    return And(*sub_results)
                elif op_type == 'or':
                    return Or(*sub_results)
        
        # 构建完整的Z3表达式
        try:
            base_expr = build_z3_expression(root_expr, expr_structure, atomic_conditions)
            print(f"构建的Z3表达式: {base_expr}")
        except Exception as e:
            print(f"构建Z3表达式出错: {str(e)}")
            traceback.print_exc()
            return []
        
        # 辅助函数
        def safe_convert_z3_value(val):
            """安全地将Z3值转换为Python值"""
            if is_int(val):
                return val.as_long()
            elif is_real(val):
                try:
                    decimal_str = val.as_decimal(10)
                    if '?' in decimal_str:
                        decimal_str = decimal_str.split('?')[0]
                    return float(decimal_str)
                except ValueError:
                    try:
                        num = val.numerator().as_long()
                        den = val.denominator().as_long()
                        if den == 0:
                            return 0.0
                        return float(num) / float(den)
                    except:
                        return 0.0
            elif is_bool(val):
                return val.__bool__()
            else:
                return float(str(val))
        
        # 为每个原子条件生成MCDC测试用例
        test_conditions = []
        for idx, (left, op, right) in enumerate(atomic_conditions):
            # 确定变量
            def get_var_name(expr):
                # 处理复合表达式
                for op_char in ['+', '-', '*', '/']:
                    if op_char in expr:
                        parts = expr.split(op_char)
                        for part in parts:
                            part = part.strip()
                            if part in z3_vars and not any(v['name'] == part and v.get('var_type') == 'constant' for v in variables):
                                return part
                
                # 普通变量
                if expr in z3_vars and not any(v['name'] == expr and v.get('var_type') == 'constant' for v in variables):
                    return expr
                return None
            
            left_var = get_var_name(left)
            right_var = get_var_name(right)
            
            # 必须至少有一个变量才能进行MCDC测试
            if left_var is None and right_var is None:
                print(f"跳过只包含常量的条件: {left} {op} {right}")
                continue
            
            # 创建当前原子条件的Z3表达式
            atom_expr = build_z3_expression(f"ATOM_{idx}", expr_structure, atomic_conditions)
            
            # 测试当前原子条件对整体表达式的影响
            # 步骤1: 找到一个情况使得原子条件为真且整体表达式为真
            s.push()
            s.add(atom_expr)  # 原子条件为真
            s.add(base_expr)  # 整体表达式为真
            
            if s.check() == sat:
                model_true = s.model()
                s.pop()
                
                # 步骤2: 找到一个情况使得原子条件为假且整体表达式为假
                s.push()
                s.add(Not(atom_expr))  # 原子条件为假
                s.add(Not(base_expr))  # 整体表达式为假
                
                if s.check() == sat:
                    model_false = s.model()
                    
                    # 创建测试用例
                    test_case_true = {}
                    test_case_false = {}
                    
                    # 填充所有变量的值
                    for var in variables:
                        name = var['name']
                        if var.get('var_type') == 'constant':
                            # 常量使用固定值
                            val = float(var['constant_value'])
                            test_case_true[name] = val
                            test_case_false[name] = val
                        else:
                            # 变量使用模型中的值
                            var_expr = z3_vars[name]
                            
                            # True case
                            if name in [str(d) for d in model_true.decls()]:
                                val = model_true.eval(var_expr)
                                test_case_true[name] = safe_convert_z3_value(val)
                            else:
                                # 默认使用中点值
                                mid = (float(var['min_value']) + float(var['max_value'])) / 2
                                if var.get('var_type') == 'int':
                                    mid = int(mid)
                                test_case_true[name] = mid
                            
                            # False case
                            if name in [str(d) for d in model_false.decls()]:
                                val = model_false.eval(var_expr)
                                test_case_false[name] = safe_convert_z3_value(val)
                            else:
                                # 默认使用与true_case相同的值
                                test_case_false[name] = test_case_true[name]
                    
                    # 添加测试用例对
                    test_conditions.append((test_case_true, test_case_false))
                    print(f"为原子条件 '{left} {op} {right}' 生成MCDC测试用例")
                
                s.pop()
            else:
                s.pop()
                print(f"无法找到使原子条件 '{left} {op} {right}' 影响整体结果的情况")
        
        print(f"生成的MCDC条件数量: {len(test_conditions)}")
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