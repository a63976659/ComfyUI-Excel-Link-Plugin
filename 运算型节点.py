import math, random, torch
import numpy as np
from . import any_typ, note



#======比较数值
class 比较数值:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入浮点数": ("FLOAT", {"default": 4.0}),
                "范围": ("STRING", {"default": "3.5-5.5"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "比较浮点数与范围"
    CATEGORY = "【Excel】联动插件/运算型节点"
    DESCRIPTION = note
    def IS_CHANGED(): return float("NaN")

    def 比较浮点数与范围(self, 输入浮点数, 范围, any=None):
        try:
            if '-' in 范围:
                下界, 上界 = map(float, 范围.split('-'))
            else:
                下界 = 上界 = float(范围)
            if 输入浮点数 < 下界:
                return ("小",)
            elif 输入浮点数 > 上界:
                return ("大",)
            else:
                return ("中",)
        except ValueError:
            return ("错误: 输入格式无效。",) 


#======规范数值
class 浮点数转整数:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "浮点数值": ("FLOAT", {"default": 3.14}),
                "操作": (["四舍五入", "取大值", "取小值", "最近32倍"], {"default": "四舍五入"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("INT",)
    FUNCTION = "转换浮点数为整数"
    CATEGORY = "【Excel】联动插件/运算型节点"
    DESCRIPTION = note
    def IS_CHANGED(): return float("NaN")

    def 转换浮点数为整数(self, 浮点数值, 操作, any=None):
        if 操作 == "四舍五入":
            结果 = round(浮点数值)
        elif 操作 == "取大值":
            结果 = math.ceil(浮点数值)
        elif 操作 == "取小值":
            结果 = math.floor(浮点数值)
        elif 操作 == "最近32倍":
            结果 = round(浮点数值 / 32) * 32
        return (结果,)


#======生成范围数组
class 生成数字:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "范围规则": ("STRING", {"default": "3|1-10"}),
                "模式": (["顺序", "随机"], {"default": "顺序"}),
                "前后缀": ("STRING", {"default": "|"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "生成数字"
    CATEGORY = "【Excel】联动插件/运算型节点"
    DESCRIPTION = note
    def IS_CHANGED(): return float("NaN")

    def 生成数字(self, 范围规则, 模式, 前后缀, any=None):
        try:
            起始字符串, 范围字符串 = 范围规则.split('|')
            起始 = int(起始字符串)
            结束范围 = list(map(int, 范围字符串.split('-')))
            if len(结束范围) == 1:
                结束 = 结束范围[0]
                数字列表 = [str(i).zfill(起始) for i in range(1, 结束 + 1)]
            else:
                起始范围, 结束 = 结束范围
                数字列表 = [str(i).zfill(起始) for i in range(起始范围, 结束 + 1)]
            if 前后缀.strip():
                前缀, 后缀 = 前后缀.split('|')
            else:
                前缀, 后缀 = "", ""
            if 模式 == "随机":
                random.shuffle(数字列表)
            数字列表 = [f"{前缀}{数字}{后缀}" for 数字 in 数字列表]
            结果 = '\n'.join(数字列表)
            return (结果,)
        except ValueError:
            return ("",)


#======范围内随机数
class 获取范围内随机整数:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "范围字符串": ("STRING", {"default": "0-10"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("INT", "STRING")
    FUNCTION = "获取范围内随机整数"
    CATEGORY = "【Excel】联动插件/运算型节点"
    DESCRIPTION = note
    def IS_CHANGED(): return float("NaN")

    def 获取范围内随机整数(self, 范围字符串, any=None):
        try:
            开始, 结束 = map(int, 范围字符串.split('-'))
            if 开始 > 结束:
                开始, 结束 = 结束, 开始
            随机整数 = random.randint(开始, 结束)
            return (随机整数, str(随机整数))
        except ValueError:
            return (0, "0")