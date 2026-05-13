#!/usr/bin/env python3
"""
Eco Agent - 文件操作与脚本执行智能体
功能：读取 .xlsx/.docx，执行 Python/R 脚本，基于 API 进行意图理解
环境：pip install openai pandas openpyxl python-docx
环境变量：setx DASHSCOPE_API_KEY "your api key"
"""

import os
import re
import sys
import json
import subprocess
from pathlib import Path
from typing import List, Dict, Any
import pandas as pd
from docx import Document
from openai import OpenAI

# ==================== 工具函数定义 ====================

def list_files(directory: str = ".") -> List[str]:
    """列出当前目录下的所有文件"""
    try:
        return [f for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]
    except Exception as e:
        return [f"错误：{e}"]

def read_xlsx(file_path: str) -> str:
    """读取 Excel 文件 (xlsx) 并返回文本内容"""
    try:
        if not os.path.exists(file_path):
            return f"错误：文件 {file_path} 不存在"
        df = pd.read_excel(file_path, sheet_name=None)
        output = []
        for sheet_name, data in df.items():
            output.append(f"工作表: {sheet_name}\n{data.to_string()}")
        return "\n\n".join(output)
    except Exception as e:
        return f"读取失败：{e}"

def read_docx(file_path: str) -> str:
    """读取 Word 文件 (docx) 并返回纯文本"""
    try:
        if not os.path.exists(file_path):
            return f"错误：文件 {file_path} 不存在"
        doc = Document(file_path)
        full_text = [para.text for para in doc.paragraphs]
        return "\n".join(full_text)
    except Exception as e:
        return f"读取失败：{e}"

def run_python(script_path: str, keywords: List[str]) -> str:
    """运行 Python 脚本，传递关键词参数"""
    try:
        if not os.path.exists(script_path):
            return f"错误：脚本 {script_path} 不存在"
        cmd = [sys.executable, script_path] + keywords
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', timeout=60)
        if result.returncode != 0:
            return f"执行失败 (code {result.returncode}):\nSTDERR:\n{result.stderr}\nSTDOUT:\n{result.stdout}"
        return result.stdout
    except subprocess.TimeoutExpired:
        return "错误：脚本执行超时（60秒）"
    except Exception as e:
        return f"执行异常：{e}"

def run_r(script_path: str, keywords: List[str]) -> str:
    """运行 R 脚本，传递关键词参数"""
    try:
        if not os.path.exists(script_path):
            return f"错误：脚本 {script_path} 不存在"
        rscript = "Rscript" if os.name != "nt" else "Rscript.exe"
        cmd = [rscript, script_path] + keywords
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            return f"执行失败 (code {result.returncode}):\nSTDERR:\n{result.stderr}\nSTDOUT:\n{result.stdout}"
        return result.stdout
    except FileNotFoundError:
        return "错误：未找到 Rscript，请确认R已安装并加入PATH"
    except subprocess.TimeoutExpired:
        return "错误：脚本执行超时（60秒）"
    except Exception as e:
        return f"执行异常：{e}"

def discover_skills(skills_root: Path) -> List[Dict[str, Any]]:
    """扫描 agentskills/*.json，提取 name, description 和 json 路径"""
    skills = []
    if not skills_root.exists():
        return skills
    for json_file in skills_root.glob("*.json"):
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                full_info = json.load(f)
            skill_basic = {
                "name": full_info.get("name", json_file.stem),
                "description": full_info.get("description", ""),
                "id": full_info.get("id", json_file.stem),
                "_json_path": json_file
            }
            skills.append(skill_basic)
        except Exception as e:
            print(f"警告：读取 {json_file} 失败：{e}")
    return skills

# ==================== 工具定义 (OpenAI function calling) ====================

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "list_files",
            "description": "列出当前工作目录下的所有文件",
            "parameters": {
                "type": "object",
                "properties": {
                    "directory": {"type": "string", "description": "目录路径，默认 '.'", "default": "."}
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "run_skill",
            "description": "执行一个已注册的技能（如'数据整理'），技能内部包含脚本和参数定义",
            "parameters": {
                "type": "object",
                "properties": {
                    "skill_name": {"type": "string", "description": "技能名称，如'数据整理'"},
                    "skill_params": {"type": "object", "description": "参数字典，键与技能定义中的参数名一致"}
                },
                "required": ["skill_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_skill_info",
            "description": "获取指定技能的详细信息，包括参数列表和工作流步骤（只读，不执行）",
            "parameters": {
                "type": "object",
                "properties": {
                    "skill_name": {"type": "string", "description": "技能名称"}
                },
                "required": ["skill_name"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_xlsx",
            "description": "读取 Excel 文件 (.xlsx) 的内容",
            "parameters": {
                "type": "object",
                "properties": {"file_path": {"type": "string", "description": ".xlsx 文件路径"}},
                "required": ["file_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "read_docx",
            "description": "读取 Word 文件 (.docx) 的内容",
            "parameters": {
                "type": "object",
                "properties": {"file_path": {"type": "string", "description": ".docx 文件路径"}},
                "required": ["file_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "run_python",
            "description": "执行 Python 脚本，传递关键词参数",
            "parameters": {
                "type": "object",
                "properties": {
                    "script_path": {"type": "string", "description": ".py 脚本路径"},
                    "keywords": {"type": "array", "items": {"type": "string"}, "description": "关键词列表"}
                },
                "required": ["script_path", "keywords"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "run_r",
            "description": "执行 R 脚本，传递关键词参数",
            "parameters": {
                "type": "object",
                "properties": {
                    "script_path": {"type": "string", "description": ".r 或 .R 脚本路径"},
                    "keywords": {"type": "array", "items": {"type": "string"}, "description": "关键词列表"}
                },
                "required": ["script_path", "keywords"]
            }
        }
    }
]

# 普通工具的函数映射
# 注意：get_skill_info ，run_skill是实例方法，不在全局映射中；将在 _execute_tool 中特殊处理
TOOL_FUNCTIONS = {
    "list_files": list_files,
    "read_xlsx": read_xlsx,
    "read_docx": read_docx,
    "run_python": run_python,
    "run_r": run_r,
}

# ==================== Agent 类 ====================

# noinspection PyTypeChecker
class EcoAnimalAgent:
    def __init__(self, api_key: str = None, max_iterations: int = 5):
        self.api_key = api_key or os.environ.get("DASHSCOPE_API_KEY")
        if not self.api_key:
            raise ValueError("请设置 DASHSCOPE_API_KEY 环境变量")
        self.client = OpenAI(
            api_key=self.api_key,
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
            timeout=60.0,
        )
        self.max_iterations = max_iterations

        # 技能发现
        self.skills_root = Path(__file__).parent / "agentskills"
        self.skills = discover_skills(self.skills_root)   # 列表，每个元素包含 name, description, id, _json_path

        # 动态生成 system prompt
        self.system_prompt = self._build_system_prompt()
        self.messages = [{"role": "system", "content": self.system_prompt}]

    def _load_full_skill(self, skill_name: str):
        """加载技能的完整配置，返回 (skill_basic, full_skill_dict) 或 (None, error_msg)"""
        skill_basic = None
        for s in self.skills:
            if s["name"] == skill_name or s.get("id") == skill_name:
                skill_basic = s
                break
        if not skill_basic:
            available = [s["name"] for s in self.skills]
            return None, f"错误：未找到技能 '{skill_name}'，可用：{', '.join(available)}"

        json_path = skill_basic["_json_path"]
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                full_skill = json.load(f)
        except Exception as e:
            return None, f"错误：无法加载技能配置 {json_path}: {e}"

        return skill_basic, full_skill

    def _get_skill_info(self, skill_name: str) -> str:
        skill_basic, full_skill = self._load_full_skill(skill_name)
        if full_skill is None:
            return skill_basic  # 此时 skill_basic 是错误消息

        info = f"技能名称: {full_skill.get('name', skill_name)}\n"
        info += f"描述: {full_skill.get('description', '无')}\n"
        params = full_skill.get("parameters", {})
        if params:
            info += "参数:\n"
            for p, desc in params.items():
                info += f"  - {p}: {desc}\n"
        else:
            info += "参数: 无\n"

        workflow = full_skill.get("workflow")
        if workflow:
            info += f"工作流步骤 (共{len(workflow)}步):\n"
            for idx, step in enumerate(workflow, 1):
                info += f"  {idx}. {step.get('description', step.get('tool', '未知操作'))}\n"
        elif full_skill.get("entry_script"):
            info += f"入口脚本: {full_skill['entry_script']}\n"
        else:
            info += "无定义的工作流或入口脚本。\n"

        return info

    def _build_system_prompt(self) -> str:
        prompt = """你是一个文件操作与脚本执行助手。可用工具：list_files, read_xlsx, read_docx, run_python, run_r, run_skill。
        当前可用的技能（仅名称和功能描述）：
        """
        if not self.skills:
            prompt += "（暂无任何技能）\n"
        else:
            for skill in self.skills:
                prompt += f"- {skill['name']}: {skill['description']}\n"
        prompt += """
        当用户请求使用某个技能时，可以先使用 get_skill_info 工具查询该技能的具体参数或工作流步骤。
        然后根据技能参数，工作流和用户的自然语言，合理组织 skill_params 字典。
        然后调用 run_skill 工具，参数为 skill_name（技能名称）和 skill_params（需要的参数）。
        用户提供的信息不足时则不传递相应参数。
        """
        return prompt

    def _run_skill(self, skill_name: str, skill_params: Dict[str, Any]) -> str:
        # 加载完整技能配置
        skill_basic, full_skill = self._load_full_skill(skill_name)
        if full_skill is None:
            return skill_basic

        # 参数校验（只校验提供的参数名是否合法）
        param_defs = full_skill.get("parameters", {})
        valid_params = set(param_defs.keys()) if param_defs else set()
        if valid_params:
            invalid_keys = set(skill_params.keys()) - valid_params
            if invalid_keys:
                return (f"错误：技能 '{skill_name}' 不支持的参数: {', '.join(invalid_keys)}\n"
                        f"支持的参数: {', '.join(valid_params)}\n"
                        f"请使用正确的参数名重新调用 run_skill。")

        # 变量替换函数
        def replace_vars(obj, params_vars):
            if isinstance(obj, str):
                result_vars = obj
                for k, v in params_vars.items():
                    result_vars = result_vars.replace(f"{{{k}}}", str(v) if v is not None else "")
                # 移除所有未被替换的占位符，如 {work_dir}
                result_vars = re.sub(r'\{[^{}]+\}', '', result_vars)
                return result_vars
            elif isinstance(obj, dict):
                return {k: replace_vars(v, params_vars) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [replace_vars(i, params_vars) for i in obj]
            else:
                return obj

        # 过滤空参数对的函数
        def filter_keywords(keywords_list):
            filtered = []
            i = 0
            while i < len(keywords_list):
                arg = keywords_list[i]
                if arg.startswith('--') and i + 1 < len(keywords_list):
                    value = keywords_list[i + 1]
                    if value is None or (isinstance(value, str) and value == ''):
                        i += 2
                        continue
                filtered.append(arg)
                i += 1
            return filtered

        # 执行工作流或回退脚本
        workflow = full_skill.get("workflow")
        if workflow:
            outputs = []
            for step_idx, action in enumerate(workflow, 1):
                tool_name = action.get("tool")
                if not tool_name:
                    outputs.append(f"步骤{step_idx} 缺少 tool 字段，跳过")
                    continue

                # 应用参数替换
                params = replace_vars(action.get("params", {}), skill_params)
                desc = action.get("description", f"步骤{step_idx}: 执行 {tool_name}")
                print(f"\n[工作流] {desc}")
                print(f"[工作流] 工具: {tool_name}, 参数: {json.dumps(params, ensure_ascii=False)}")

                # 调用工具函数
                result = None
                if tool_name == "run_skill":
                    result = "错误：工作流中不能嵌套调用 run_skill"
                elif tool_name == "get_skill_info":
                    result = "错误：工作流中不能调用 get_skill_info"
                else:
                    tool_func = TOOL_FUNCTIONS.get(tool_name)
                    if not tool_func:
                        result = f"错误：未知工具 {tool_name}"
                    else:
                        try:
                            # 特殊处理 run_python 和 run_r：过滤空参数
                            if tool_name in ("run_python", "run_r") and "keywords" in params:
                                params["keywords"] = filter_keywords(params["keywords"])
                            result = tool_func(**params)
                        except Exception as e:
                            result = f"执行异常：{e}"

                # 记录输出
                output_preview = result
                outputs.append(f"【步骤{step_idx}】{desc}\n结果: {output_preview}")

                if action.get("continue_on_error") is False and ("错误" in result or "失败" in result):
                    outputs.append("工作流因错误终止。")
                    break

            return "\n\n".join(outputs)
        else:
            # 向后兼容：使用 entry_script
            entry_script = full_skill.get("entry_script")
            if not entry_script:
                return "错误：技能未定义 workflow 或 entry_script"
            script_path = skill_basic["_json_path"].parent / entry_script
            if not script_path.exists():
                return f"错误：技能脚本不存在 {script_path}"

            cmd = [sys.executable, str(script_path)]
            for param_name, param_value in skill_params.items():
                cmd.extend([f"--{param_name}", str(param_value)])

            try:
                result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', timeout=60)
                if result.returncode != 0:
                    return f"执行失败 (code {result.returncode}):\nSTDERR:\n{result.stderr}\nSTDOUT:\n{result.stdout}"
                return result.stdout
            except subprocess.TimeoutExpired:
                return "错误：脚本执行超时（60秒）"
            except Exception as e:
                return f"执行异常：{e}"

    def _execute_tool(self, tool_call) -> Dict[str, Any]:
        func_name = tool_call.function.name
        args = json.loads(tool_call.function.arguments)
        print(f"\n[Tool] 调用工具: {func_name}")
        print(f"[Tool] 参数: {json.dumps(args, ensure_ascii=False, indent=2)}")

        if func_name == "run_skill":
            result = self._run_skill(**args)
        elif func_name == "get_skill_info":
            result = self._get_skill_info(**args)
        else:
            func = TOOL_FUNCTIONS.get(func_name)
            if not func:
                result = f"错误：未知工具 {func_name}"
            else:
                result = func(**args)

        # 显示结果
        if isinstance(result, str) and len(result) > 500:
            display_result = result[:500] + "...(结果过长，已截断)"
        else:
            display_result = result
        print(f"[Tool] 执行结果:\n{display_result}\n")

        # 存储时截断
        if isinstance(result, str) and len(result) > 3000:
            result = result[:3000] + "...(内容过长已截断)"
        return {"result": result}

    # noinspection PyTypeHints
    def run(self, user_input: str) -> str:
        self.messages.append({"role": "user", "content": user_input})
        print(f"\n[User] {user_input}")

        iteration = 0
        while iteration < self.max_iterations:
            # 打印即将发送给 LLM 的消息（简化版）
            print("-" * 50)
            print(f"\n[助手] 准备请求 (第{iteration + 1}轮)，当前消息数: {len(self.messages)}")
            # 可选：打印最近几条消息的内容，避免刷屏
            for idx, msg in enumerate(self.messages[-3:]):  # 只打印最后条
                role = msg.get("role", "unknown")
                content = msg.get("content", "")
                if content and len(content) > 500:
                    content = content[:500] + "..."
                print(f"  - {role}: {content if content else '(无文本)'}")
                if msg.get("tool_calls"):
                    print(f"    tool_calls: {[tc['function']['name'] for tc in msg['tool_calls']]}")

            try:
                response = self.client.chat.completions.create(
                    model="qwen-plus",
                    messages=self.messages,
                    tools=TOOLS,
                    tool_choice="auto",
                    temperature=0.1,
                    max_tokens=2000,
                )
            except Exception as e:
                error_msg = f"API 调用失败: {e}"
                print(f"[Error] {error_msg}")
                return error_msg

            message = response.choices[0].message
            # 打印 LLM 响应
            if message.content:
                print(f"\n[LLM] 回复文本: {message.content}")
            if message.tool_calls:
                print(f"[LLM] 请求调用 {len(message.tool_calls)} 个工具:")
                for tc in message.tool_calls:
                    print(f"  - {tc.function.name}({tc.function.arguments})")

            self.messages.append(message.model_dump(exclude_unset=True))

            if not message.tool_calls:
                final_answer = message.content or "（模型无文本返回）"
                return final_answer

            for tool_call in message.tool_calls:
                tool_result = self._execute_tool(tool_call)
                self.messages.append({
                    "role": "tool",
                    "tool_call_id": tool_call.id,
                    "content": json.dumps(tool_result, ensure_ascii=False)
                })
            iteration += 1

        error = "达到最大工具调用次数，无法完成请求。请简化指令。"
        print(f"[Error] {error}")
        return error

    def chat_loop(self):
        print("EcoAnimalAgent 已启动。输入 'exit' 或 'quit' 退出。")
        print("示例请求：")
        print("  列出当前文件")
        print("  读取 report.docx")
        print("  运行 analysis.py，关键词为 'sales, profit'")
        print("  执行 model.R，传递关键词 '2024', 'prediction'")
        print("-" * 50)
        while True:
            try:
                user_input = input("\n> ").strip()
                if user_input.lower() in ("exit", "quit"):
                    print("再见！")
                    break
                if not user_input:
                    continue
                response = self.run(user_input)
                print(f"\n助手: {response}")
            except KeyboardInterrupt:
                print("\n程序中断")
                break
            except Exception as e:
                print(f"\n错误: {e}")

# ==================== 命令行入口 ====================
if __name__ == "__main__":
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_dir)
    except NameError:
        pass
    agent = EcoAnimalAgent()
    agent.chat_loop()