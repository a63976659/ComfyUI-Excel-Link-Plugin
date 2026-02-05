import os, time, secrets, requests, random, base64, openpyxl, datetime
import folder_paths
import numpy as np
from PIL import Image
from . import any_typ, note

#======当前时间(戳)
class 获取当前时间:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "前缀": ("STRING", {"default": ""}),
            },
            "optional": {"任意": (any_typ,)} 
        }
    
    RETURN_TYPES = ("STRING", any_typ)
    RETURN_NAMES = ("时间文本", "任意输出")
    FUNCTION = "获取时间"
    CATEGORY = "【Excel】联动插件/功能节点"
    DESCRIPTION = "【使用方法】获取当前系统时间。输出标准的文本格式（YYYY-MM-DD HH:MM:SS），可直接连接至写入节点。"
    OUTPUT_NODE = True
    
    def IS_CHANGED(self, **kwargs):
        return float("NaN")

    def 获取时间(self, 前缀, any=None):
        try:
            now = datetime.datetime.now()
            # 格式化为标准的秒级文本
            time_str = now.strftime("%Y-%m-%d %H:%M:%S")
            
            res_str = f"{前缀} {time_str}".strip() if 前缀 else time_str
            
            return (res_str, any)
        except Exception as e:
            return (f"时间获取失败: {str(e)}", any)

#======随机整数
class 简单随机种子:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING", "INT")
    RETURN_NAMES = ("种子文本", "种子数值")
    FUNCTION = "生成随机种子"
    CATEGORY = "【Excel】联动插件/功能节点"
    DESCRIPTION = "【使用方法】生成一个高位随机整数作为随机种子。输出提供文本和数值两种类型。"
    OUTPUT_NODE = True
    
    def IS_CHANGED(self, any=None):
        return float("NaN")

    def 生成随机种子(self, any=None):
        seed = random.randint(100000000, 999999999)
        return (str(seed), seed)

#======选择参数
class 选择参数:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "性别": (["男性", "女性"], {"default": "男性"}),
                "版本": (["竖版", "横版"], {"default": "竖版"}),
                "附加文本": ("STRING", {"multiline": True, "default": ""}),
            },
            "optional": {"任意": (any_typ,)} 
        }
    
    RETURN_TYPES = ("STRING",)
    RETURN_NAMES = ("组合结果",)
    FUNCTION = "执行选择"
    CATEGORY = "【Excel】联动插件/功能节点"
    DESCRIPTION = "【使用方法】根据性别和画幅版本生成分类标识（如 1+1），并自动拼接填写的附加文本。"
    OUTPUT_NODE = True
    
    def IS_CHANGED(self, **kwargs):
        return float("NaN")

    def 执行选择(self, 性别, 版本, 附加文本, any=None):
        s_val = 1 if 性别 == "男性" else 2
        v_val = 1 if 版本 == "竖版" else 2
        res = f"{s_val}+{v_val}\n\n{附加文本.strip()}"
        return (res,)

#======读取页面
class 读取网页节点:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "指令": ("STRING", {"default": ""}),
                "前后缀": ("STRING", {"default": ""}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING",)
    RETURN_NAMES = ("网页内容",)
    FUNCTION = "获取数据"
    CATEGORY = "【Excel】联动插件/功能节点"
    DESCRIPTION = "【使用方法】通过特定接口指令获取远程网页的文本内容。"
    OUTPUT_NODE = True
    
    def IS_CHANGED(self, **kwargs):
        return float("NaN")

    def 获取数据(self, 指令, 前后缀, any=None):
        try:
            base_url = base64.b64decode('aHR0cHM6Ly93d3cubWVlZXlvLmNvbS91L2dldG5vZGUv').decode()
            ext = base64.b64decode('LnBocA==').decode()
            target_url = f"{base_url}{指令.lower()}{ext}"

            token = secrets.token_hex(16)
            headers = {'Authorization': f'Bearer {token}'}
            response = requests.get(target_url, headers=headers, timeout=10)
            response.raise_for_status()
            
            prefix, suffix = 前后缀.split("|", 1) if "|" in 前后缀 else (前后缀, "")
            final_text = f"{prefix}{response.text}{suffix}"
            return (final_text,)
        except Exception as e:
            return (f"网页读取失败: {str(e)}",)

#======完成提醒
class 完成提醒:
    def __init__(self):
        self.audio_files = self._get_audio_list()
    
    def _get_audio_list(self):
        try:
            curr_dir = os.path.dirname(os.path.abspath(__file__))
            audio_dir = os.path.join(curr_dir, "音频")
            if not os.path.exists(audio_dir):
                return ["notify.mp3"]
            files = [f for f in os.listdir(audio_dir) if f.lower().endswith(('.mp3', '.wav'))]
            return sorted(files) if files else ["notify.mp3"]
        except:
            return ["notify.mp3"]
    
    @classmethod
    def INPUT_TYPES(cls):
        instance = cls()
        return {
            "required": {
                "模式": (["总是", "空列队"], {"default": "总是"}),
                "音量": ("FLOAT", {"min": 0, "max": 100, "step": 1, "default": 50}),
                "音频文件": (instance.audio_files, {"default": instance.audio_files[0]}),
            },
            "optional": {"任意": (any_typ,)}
        }

    RETURN_TYPES = (any_typ,)
    RETURN_NAMES = ("任意输出",)
    FUNCTION = "执行提醒"
    CATEGORY = "【Excel】联动插件/功能节点"
    DESCRIPTION = "【使用方法】当流程运行至此时播放音频。"
    OUTPUT_NODE = True

    def IS_CHANGED(self, **kwargs):
        return float("NaN")

    def 执行提醒(self, 模式, 音量, 音频文件, 任意=None):
        try:
            curr_dir = os.path.dirname(os.path.abspath(__file__))
            audio_path = os.path.join(curr_dir, "音频", 音频文件)
            if os.path.exists(audio_path):
                if os.name == 'nt': 
                    os.startfile(audio_path)
                else: 
                    os.system(f"open '{audio_path}' || xdg-open '{audio_path}'")
            return (任意 if 任意 is not None else "Done",)
        except:
            return ("Error",)