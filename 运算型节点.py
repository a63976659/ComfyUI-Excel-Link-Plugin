import math
import random
import torch
import numpy as np
from . import any_typ, note

#====== 比较数值
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
    DESCRIPTION = "【使用方法】输入一个数值和范围（如 3.5-5.5）。如果数值小于下限返回'小'，大于上限返回'大'，在范围内返回'中'。常用于逻辑判断或条件分支。"
    
    def IS_CHANGED(self, **kwargs): return float("NaN")

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
            return ("错误: 范围格式无效 (应为 0.0-1.0)。",) 

#====== 规范数值
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
    DESCRIPTION = "【使用方法】对输入的浮点数进行取整处理。'最近32倍'选项特别适用于 Stable Diffusion 的尺寸规范要求（图像长宽通常需为32或64的倍数）。"
    
    def IS_CHANGED(self, **kwargs): return float("NaN")

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

#====== 生成范围数组
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
    FUNCTION = "执行生成"
    CATEGORY = "【Excel】联动插件/运算型节点"
    DESCRIPTION = "【使用方法】按规则生成数字序列。'范围规则'格式：'位数|起始-结束'（如 3|1-10 生成 001-010）。'前后缀'用'|'分隔前缀和后缀。常用于批量生成文件名。"
    
    def IS_CHANGED(self, **kwargs): return float("NaN")

    def 执行生成(self, 范围规则, 模式, 前后缀, any=None):
        try:
            起始字符串, 范围字符串 = 范围规则.split('|')
            zfill_num = int(起始字符串)
            
            if '-' in 范围字符串:
                start_range, end_range = map(int, 范围字符串.split('-'))
            else:
                start_range, end_range = 1, int(范围字符串)
                
            num_list = [str(i).zfill(zfill_num) for i in range(start_range, end_range + 1)]
            
            if 模式 == "随机":
                random.shuffle(num_list)
            
            pre, suf = 前后缀.split('|') if '|' in 前后缀 else (前后缀, "")
            res_list = [f"{pre}{n}{suf}" for n in num_list]
            
            return ('\n'.join(res_list),)
        except Exception as e:
            return (f"生成失败: {str(e)}",)

#====== 范围内随机数
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
    RETURN_NAMES = ("整数值", "文本格式")
    FUNCTION = "获取随机整数"
    CATEGORY = "【Excel】联动插件/运算型节点"
    DESCRIPTION = "【使用方法】从给定范围（如 1-100）中随机抽取一个整数。同时输出整数类型和字符串类型，方便连接到不同类型的节点输入端。"
    
    def IS_CHANGED(self, **kwargs): return float("NaN")

    def 获取随机整数(self, 范围字符串, any=None):
        try:
            start, end = map(int, 范围字符串.split('-'))
            if start > end:
                start, end = end, start
            res = random.randint(start, end)
            return (res, str(res))
        except ValueError:
            return (0, "0")
