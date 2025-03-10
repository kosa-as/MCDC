import os
import subprocess
import sys
import time

def run_script(script_name):
    """
    运行指定的Python脚本
    
    Args:
        script_name (str): 脚本文件名
    
    Returns:
        bool: 脚本是否成功执行
    """
    print(f"\n{'='*80}")
    print(f"正在执行 {script_name}...")
    print(f"{'='*80}\n")
    
    try:
        # 使用当前Python解释器执行脚本
        result = subprocess.run([sys.executable, script_name], check=True)
        print(f"\n{script_name} 执行成功！")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n错误: {script_name} 执行失败，返回码: {e.returncode}")
        return False
    except Exception as e:
        print(f"\n错误: 执行 {script_name} 时发生异常: {e}")
        return False

def check_required_files():
    """
    检查必要的输入文件是否存在
    
    Returns:
        bool: 所有必要文件是否存在
    """
    required_files = [
        "input/testcase.xlsx",
        "input/module.docx",
        "input/data.docx"
    ]
    
    missing_files = []
    for file_path in required_files:
        if not os.path.exists(file_path):
            missing_files.append(file_path)
    
    if missing_files:
        print("错误: 以下必要文件不存在:")
        for file_path in missing_files:
            print(f"  - {file_path}")
        return False
    
    return True

def check_output_files():
    """
    检查处理过程中生成的中间文件是否存在
    
    Returns:
        bool: 所有中间文件是否存在
    """
    required_files = [
        "output/testcase_parsed.json",
        "output/testcase_parsed_merged.json",
        "output/module_modules.json",
        "output/testcase_parsed_merged_with_results.json",
        "output/data_constants.json"
    ]
    
    missing_files = []
    for file_path in required_files:
        if not os.path.exists(file_path):
            missing_files.append(file_path)
    
    if missing_files:
        print("警告: 以下中间文件不存在:")
        for file_path in missing_files:
            print(f"  - {file_path}")
        return False
    
    return True

def main():
    """主函数，按顺序执行所有处理步骤"""
    start_time = time.time()
    
    # 确保输出目录存在
    os.makedirs("output", exist_ok=True)
    
    # 检查必要的输入文件
    if not check_required_files():
        print("处理终止: 缺少必要的输入文件")
        return
    
    # 执行testcase_parser.py
    if not run_script("testcase_parser.py"):
        print("处理终止: testcase_parser.py 执行失败")
        return
    
    # 检查中间文件是否生成
    if not check_output_files():
        print("警告: 部分中间文件未生成，但将继续执行")
    
    # 执行generate_testcase.py
    if not run_script("generate_testcase.py"):
        print("处理终止: generate_testcase.py 执行失败")
        return
    
    # 检查最终输出文件
    final_output = "output/generated_testcases.xlsx"
    if os.path.exists(final_output):
        print(f"\n处理完成！最终结果已保存至: {final_output}")
    else:
        print(f"\n警告: 最终输出文件 {final_output} 未生成")
    
    # 计算总耗时
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"总耗时: {elapsed_time:.2f} 秒")

if __name__ == "__main__":
    print("开始执行测试用例处理流程...")
    main()
    print("\n处理流程结束") 