class Variable:
    def __init__(self, name, symbol, var_type, type_desc, initial_value, 
                 comment, identifier, min_value, max_value):
        self.name = name                # 变量名
        self.symbol = symbol            # 变量符号
        self.var_type = var_type        # 类型
        self.type_desc = type_desc      # 类型解释
        self.initial_value = initial_value  # 初始值
        self.comment = comment          # 变量注释
        self.identifier = identifier    # 变量标识位
        self.min_value = min_value      # 最小值
        self.max_value = max_value      # 最大值

class Module:
    def __init__(self, name):
        self.name = name            # 任务名称
        self.number = ""           # 编号
        self.function = ""         # 功能
        self.precondition = ""     # 前置条件
        self.inputs = []           # 输入
        self.outputs = []          # 输出
        self.formula = ""          # 公式

class TestDataManager:
    def __init__(self):
        self.variables = {}   # 存储变量
        self.constants = {}   # 存储常量
        self.modules = {}     # 存储模块