import re, math, datetime, random, secrets, requests, string
from . import any_typ, note



#======文本输入
class 文本输入:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "文本": ("STRING", {"default": "", "multiline": True}),
            }
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "处理输入"
    OUTPUT_NODE = False
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 处理输入(self, 文本):
        return (文本,)


#======文本到列表
class 文本到列表:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "文本": ("STRING", {"multiline": True, "default": ""}),
                "分隔符": ("STRING", {"default": ""}),
            }
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "分割文本"
    OUTPUT_IS_LIST = (True,)
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 分割文本(self, 文本, 分隔符):
        if not 分隔符:
            部分 = 文本.split('\n')
        else:
            部分 = 文本.split(分隔符)
        部分 = [部分.strip() for 部分 in 部分 if 部分.strip()]
        if not 部分:
            return ([],)
        return (部分,)


#======文本拼接
class 文本拼接器:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "optional": {
                "文本1": ("STRING", {"multiline": False, "default": ""}),
                "文本2": ("STRING", {"multiline": False, "default": ""}),
                "文本3": ("STRING", {"multiline": False, "default": ""}),
                "文本4": ("STRING", {"multiline": False, "default": ""}),
                "组合顺序": ("STRING", {"default": ""}),
                "分隔符": ("STRING", {"default": ","})
            },
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "组合文本"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 组合文本(self, 文本1, 文本2, 文本3, 文本4, 组合顺序, 分隔符):
        try:
            文本映射 = {
                "1": 文本1,
                "2": 文本2,
                "3": 文本3,
                "4": 文本4
            }
            if not 组合顺序:
                组合顺序 = "1+2+3+4"
            部分 = 组合顺序.split("+")
            有效部分 = []
            for 部分 in 部分:
                if 部分 in 文本映射:
                    有效部分.append(部分)
                else:
                    return (f"错误: 组合顺序中的输入 '{部分}' 无效。有效选项为 1, 2, 3, 4。",)
            非空文本 = [文本映射[部分] for 部分 in 有效部分 if 文本映射[部分]]
            
            if 分隔符 == '\\n':
                分隔符 = '\n'
            
            结果 = 分隔符.join(非空文本) 
            return (结果,)
        except Exception as e:
            return (f"错误: {str(e)}",)
        

#======多参数输入
class 多参数输入节点:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "文本1": ("STRING", {"default": "", "multiline": True}),  
                "文本2": ("STRING", {"default": "", "multiline": True}), 
                "整数1": ("INT", {"default": 0, "min": -1000000, "max": 1000000}),  
                "整数2": ("INT", {"default": 0, "min": -1000000, "max": 1000000}), 
            }
        }

    RETURN_TYPES = ("STRING", "STRING", "INT", "INT")
    FUNCTION = "处理输入"
    OUTPUT_NODE = False
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 处理输入(self, 文本1, 文本2, 整数1, 整数2):
        return (文本1, 文本2, 整数1, 整数2)


#======整数参数
class 数字提取器:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"default": "2|3"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("INT", "INT")
    FUNCTION = "按索引提取行"
    OUTPUT_TYPES = ("INT", "INT")
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 按索引提取行(self, 输入文本):
        try:
            数据列表 = 输入文本.split("|")
            
            结果 = []
            for i in range(2): 
                if i < len(数据列表):
                    try:
                        结果.append(int(数据列表[i]))
                    except ValueError:
                        结果.append(0)
                else:
                    结果.append(0)
            
            return tuple(结果)
        except:
            return (0, 0)


#======添加前后缀
class 添加前后缀:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"default": ""}),
                "前缀": ("STRING", {"default": "前缀符"}),
                "后缀": ("STRING", {"default": "后缀符"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "添加前后缀"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 添加前后缀(self, 输入字符串, 前缀, 后缀):
        return (f"{前缀}{输入字符串}{后缀}",)


#======提取标签之间
class 提取子字符串:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"multiline": True, "default": ""}),  
                "模式": ("STRING", {"default": "标签始|标签尾"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "提取子字符串"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 提取子字符串(self, 输入字符串, 模式):
        try:
            部分 = 模式.split('|')
            开始字符串 = 部分[0]
            结束字符串 = 部分[1] if len(部分) > 1 and 部分[1].strip() else "\n"

            开始索引 = 输入字符串.index(开始字符串) + len(开始字符串)

            结束索引 = 输入字符串.find(结束字符串, 开始索引)
            if 结束索引 == -1:
                结束索引 = 输入字符串.find("\n", 开始索引)
                if 结束索引 == -1:
                    结束索引 = len(输入字符串)

            提取内容 = 输入字符串[开始索引:结束索引]

            行列表 = 提取内容.splitlines()
            非空行 = [行 for 行 in 行列表 if 行.strip()]
            结果 = '\n'.join(非空行)

            return (结果,)
        except ValueError:
            return ("",)


#======按数字范围提取
class 按索引提取子字符串:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"default": ""}),
                "索引范围": ("STRING", {"default": "2-6"}),
                "方向": (["从前面", "从后面"], {"default": "从前面"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "按索引提取子字符串"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 按索引提取子字符串(self, 输入字符串, 索引范围, 方向):
        try:
            if '-' in 索引范围:
                开始索引, 结束索引 = map(int, 索引范围.split('-'))
            else:
                开始索引 = 结束索引 = int(索引范围)

            开始索引 -= 1
            结束索引 -= 1

            if 开始索引 < 0 or 开始索引 >= len(输入字符串):
                return ("",)

            if 结束索引 < 0 or 结束索引 >= len(输入字符串):
                结束索引 = len(输入字符串) - 1

            if 开始索引 > 结束索引:
                开始索引, 结束索引 = 结束索引, 开始索引

            if 方向 == "从前面":
                return (输入字符串[开始索引:结束索引 + 1],)
            elif 方向 == "从后面":
                开始索引 = len(输入字符串) - 开始索引 - 1
                结束索引 = len(输入字符串) - 结束索引 - 1
                if 开始索引 > 结束索引:
                    开始索引, 结束索引 = 结束索引, 开始索引
                return (输入字符串[开始索引:结束索引 + 1],)
        except ValueError:
            return ("",)
			

#======分隔符拆分两边
class 按分隔符拆分字符串:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"default": "文本|内容"}),
                "分隔符": ("STRING", {"default": "|"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING", "STRING")
    FUNCTION = "按分隔符拆分字符串"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 按分隔符拆分字符串(self, 输入字符串, 分隔符):
        部分 = 输入字符串.split(分隔符, 1)
        if len(部分) == 2:
            return (部分[0], 部分[1])
        else:
            return (输入字符串, "")


#======常规处理字符
class 处理字符串:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"multiline": True, "default": ""}),
                "选项": (["不改变", "取数字", "取字母", "转大写", "转小写", "取中文", "去标点", "去换行", "去空行", "去空格", "去格式", "统计字数"], {"default": "不改变"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "处理字符串"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 处理字符串(self, 输入字符串, 选项):
        if 选项 == "取数字":
            结果 = ''.join(re.findall(r'\d', 输入字符串))
        elif 选项 == "取字母":
            结果 = ''.join(filter(lambda 字符: 字符.isalpha() and not self.是中文字符(字符), 输入字符串))
        elif 选项 == "转大写":
            结果 = 输入字符串.upper()
        elif 选项 == "转小写":
            结果 = 输入字符串.lower()
        elif 选项 == "取中文":
            结果 = ''.join(filter(self.是中文字符, 输入字符串))
        elif 选项 == "去标点":
            结果 = re.sub(r'[^\w\s\u4e00-\u9fff]', '', 输入字符串)
        elif 选项 == "去换行":
            结果 = 输入字符串.replace('\n', '')
        elif 选项 == "去空行":
            结果 = '\n'.join(filter(lambda 行: 行.strip(), 输入字符串.splitlines()))
        elif 选项 == "去空格":
            结果 = 输入字符串.replace(' ', '')
        elif 选项 == "去格式":
            结果 = re.sub(r'\s+', '', 输入字符串)
        elif 选项 == "统计字数":
            结果 = str(len(输入字符串))
        else:
            结果 = 输入字符串

        return (结果,)

    @staticmethod
    def 是中文字符(字符):
        return '\u4e00' <= 字符 <= '\u9fff'


#======提取前后字符
class 提取前后字符:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"default": ""}),
                "模式": ("STRING", {"default": "标签符"}),
                "位置": (["保留最初之前", "保留最初之后", "保留最后之前", "保留最后之后"], {"default": "保留最初之前"}),
                "包含分隔符": ("BOOLEAN", {"default": False}), 
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "提取前后字符"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 提取前后字符(self, 输入字符串, 模式, 位置, 包含分隔符):
        if 位置 == "保留最初之前":
            索引 = 输入字符串.find(模式)
            if 索引 != -1:
                结果 = 输入字符串[:索引 + len(模式) if 包含分隔符 else 索引]
                return (结果,)
        elif 位置 == "保留最初之后":
            索引 = 输入字符串.find(模式)
            if 索引 != -1:
                结果 = 输入字符串[索引:] if 包含分隔符 else 输入字符串[索引 + len(模式):]
                return (结果,)
        elif 位置 == "保留最后之前":
            索引 = 输入字符串.rfind(模式)
            if 索引 != -1:
                结果 = 输入字符串[:索引 + len(模式) if 包含分隔符 else 索引]
                return (结果,)
        elif 位置 == "保留最后之后":
            索引 = 输入字符串.rfind(模式)
            if 索引 != -1:
                结果 = 输入字符串[索引:] if 包含分隔符 else 输入字符串[索引 + len(模式):]
                return (结果,)
        return ("",)


#======简易文本替换
class 简易文本替换:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"multiline": True, "default": "", "forceInput": True}),
                "查找文本": ("STRING", {"default": ""}),
                "替换文本": ("STRING", {"default": ""})
            },
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "替换文本"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 替换文本(self, 输入字符串, 查找文本, 替换文本):
        try:
            if not 查找文本:
                return (输入字符串,)

            if 替换文本 == '\\n':
                替换文本 = '\n'
            
            结果 = 输入字符串.replace(查找文本, 替换文本)
            return (结果,)
        except Exception as e:
            return (f"错误: {str(e)}",)
        

#======替换第n次出现
class 替换第n次出现:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "原始文本": ("STRING", {"multiline": True, "default": ""}),
                "出现次数": ("INT", {"default": 1, "min": 0}),
                "查找字符串": ("STRING", {"default": "替换前字符"}),
                "替换字符串": ("STRING", {"default": "替换后字符"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "替换第n次出现"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 替换第n次出现(self, 原始文本, 出现次数, 查找字符串, 替换字符串):
        if 出现次数 == 0:
            结果 = 原始文本.replace(查找字符串, 替换字符串)
        else:
            def 替换第n次匹配(匹配):
                nonlocal 出现次数
                出现次数 -= 1
                return 替换字符串 if 出现次数 == 0 else 匹配.group(0)

            结果 = re.sub(re.escape(查找字符串), 替换第n次匹配, 原始文本, count=出现次数)

        return (结果,)


#======多次出现依次替换
class 多次替换:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "原始文本": ("STRING", {"multiline": True, "default": ""}),
                "替换规则": ("STRING", {"default": ""}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "多次替换"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 多次替换(self, 原始文本, 替换规则):
        try:
            搜索字符串, 替换列表 = 替换规则.split('|')
            替换列表 = [替换 for 替换 in 替换列表.split(',') if 替换]

            def 替换匹配(匹配):
                nonlocal 替换列表
                if 替换列表:
                    return 替换列表.pop(0)
                return 匹配.group(0)

            结果 = re.sub(re.escape(搜索字符串), 替换匹配, 原始文本)

            return (结果,)
        except ValueError:
            return ("",)


#======批量替换字符
class 批量替换字符串:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "原始文本": ("STRING", {"multiline": False, "default": "文本内容"}),
                "替换规则": ("STRING", {"multiline": True, "default": ""}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "批量替换字符串"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 批量替换字符串(self, 原始文本, 替换规则):
        规则列表 = 替换规则.strip().split('\n')
        for 规则 in 规则列表:
            if '|' in 规则:
                搜索字符串列表, 替换字符串 = 规则.split('|', 1)
                
                搜索字符串列表 = 搜索字符串列表.replace("\\n", "\n")
                替换字符串 = 替换字符串.replace("\\n", "\n")
                
                搜索字符串列表 = 搜索字符串列表.split(',')
                
                for 搜索字符串 in 搜索字符串列表:
                    原始文本 = 原始文本.replace(搜索字符串, 替换字符串)
        return (原始文本,)


#======随机行内容
class 从文本随机行:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "获取随机行"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 获取随机行(self, 输入文本, any=None):
        行列表 = 输入文本.strip().splitlines()
        if not 行列表:
            return ("",)  
        return (random.choice(行列表),)


#======判断是否包含字符
class 判断是否包含字符:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"default": "文本内容"}),
                "子字符串": ("STRING", {"default": "查找符1|查找符2"}),
                "模式": (["同时满足", "任意满足"], {"default": "任意满足"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("INT",)
    FUNCTION = "检查子字符串存在"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 检查子字符串存在(self, 输入文本, 子字符串, 模式):
        子字符串列表 = 子字符串.split('|')

        if 模式 == "同时满足":
            for 子字符串 in 子字符串列表:
                if 子字符串 not in 输入文本:
                    return (0,)
            return (1,)
        elif 模式 == "任意满足":
            for 子字符串 in 子字符串列表:
                if 子字符串 in 输入文本:
                    return (1,)
            return (0,)


#======段落每行添加前后缀
class 每行添加前后缀:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}),  
                "前后缀": ("STRING", {"default": "前缀符|后缀符"}),  
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "每行添加前后缀"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 每行添加前后缀(self, 前后缀, 输入文本):
        try:
            前缀, 后缀 = 前后缀.split('|')
            行列表 = 输入文本.splitlines()
            修改后的行 = [f"{前缀}{行}{后缀}" for 行 in 行列表]
            结果 = '\n'.join(修改后的行)
            return (结果,)
        except ValueError:
            return ("",)  


#======段落提取指定索引行
class 段落提取指定索引行:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "行索引": ("STRING", {"default": "2-3"}), 
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "段落提取指定索引行"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 段落提取指定索引行(self, 输入文本, 行索引):
        try:
            行列表 = 输入文本.splitlines()
            结果行 = []

            if '-' in 行索引:
                开始, 结束 = map(int, 行索引.split('-'))
                开始 = max(1, 开始)  
                结束 = min(len(行列表), 结束)  
                结果行 = 行列表[开始 - 1:结束]
            else:
                索引列表 = map(int, 行索引.split('|'))
                for 索引 in 索引列表:
                    if 1 <= 索引 <= len(行列表):
                        结果行.append(行列表[索引 - 1])

            结果 = '\n'.join(结果行)
            return (结果,)
        except ValueError:
            return ("",) 


#======段落提取或移除字符行
class 段落提取或移除字符行:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "子字符串": ("STRING", {"default": "查找符1|查找符2"}), 
                "操作": (["保留", "移除"], {"default": "保留"}), 
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "段落提取或移除字符行"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 段落提取或移除字符行(self, 输入文本, 子字符串, 操作):
        行列表 = 输入文本.splitlines()
        子字符串列表 = 子字符串.split('|')
        结果行 = []

        for 行 in 行列表:
            包含子字符串 = any(子字符串 in 行 for 子字符串 in 子字符串列表)
            if (操作 == "保留" and 包含子字符串) or (操作 == "移除" and not 包含子字符串):
                结果行.append(行)

        非空行 = [行 for 行 in 结果行 if 行.strip()]
        结果 = '\n'.join(非空行)
        return (结果,)


#======根据字数范围过滤文本行
class 按字数过滤行:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "字数范围": ("STRING", {"default": "2-10"}),  
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "按字数过滤行"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 按字数过滤行(self, 输入文本, 字数范围):
        try:
            行列表 = 输入文本.splitlines()
            结果行 = []

            if '-' in 字数范围:
                最小字数, 最大字数 = map(int, 字数范围.split('-'))
                结果行 = [行 for 行 in 行列表 if 最小字数 <= len(行) <= 最大字数]

            结果 = '\n'.join(结果行)
            return (结果,)
        except ValueError:
            return ("",)  


#======按序号提取分割文本
class 分割并提取文本:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}),
                "分隔符": ("STRING", {"default": "分隔符"}),
                "索引": ("INT", {"default": 1, "min": 1}),
                "顺序": (["顺序", "倒序"], {"default": "顺序"}),
                "包含分隔符": ("BOOLEAN", {"default": False}), 
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "分割并提取"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 分割并提取(self, 输入文本, 分隔符, 索引, 顺序, 包含分隔符):
        try:
            if not 分隔符:
                部分 = 输入文本.splitlines()
            else:
                部分 = 输入文本.split(分隔符)
            
            if 顺序 == "倒序":
                部分 = 部分[::-1]
            
            if 索引 > 0 and 索引 <= len(部分):
                选中部分 = 部分[索引 - 1]
                
                if 包含分隔符 and 分隔符:
                    if 顺序 == "顺序":
                        if 索引 > 1:
                            选中部分 = 分隔符 + 选中部分
                        if 索引 < len(部分):
                            选中部分 += 分隔符
                    elif 顺序 == "倒序":
                        if 索引 > 1:
                            选中部分 += 分隔符
                        if 索引 < len(部分):
                            选中部分 = 分隔符 + 选中部分
                
                行列表 = 选中部分.splitlines()
                非空行 = [行 for 行 in 行列表 if 行.strip()]
                结果 = '\n'.join(非空行)
                return (结果,)
            else:
                return ("",)
        except ValueError:
            return ("",)


#======文本出现次数
class 文本出现次数:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}),
                "字符": ("STRING", {"default": "查找符"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("INT", "STRING")
    FUNCTION = "文本出现次数"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 文本出现次数(self, 输入文本, 字符):
        try:
            if 字符 == "\\n":
                行列表 = [行 for 行 in 输入文本.splitlines() if 行.strip()]
                计数 = len(行列表)
            else:
                计数 = 输入文本.count(字符)
            return (计数, str(计数))
        except ValueError:
            return (0, "0")


#======文本拆分
class 文本拆分:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}),
                "分隔符": ("STRING", {"default": "标签符"}),  
                "索引": ("INT", {"default": 1, "min": 1}),  
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING", "STRING", "STRING", "STRING", "STRING")
    FUNCTION = "文本拆分"
    OUTPUT_TYPES = ("STRING", "STRING", "STRING", "STRING", "STRING")
    OUTPUT_NAMES = ("文本1", "文本2", "文本3", "文本4", "文本5")
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 文本拆分(self, 输入文本, 索引, 分隔符):
        try:
            if 分隔符 == "" or 分隔符 == "\n":
                行列表 = 输入文本.splitlines()
            else:
                行列表 = 输入文本.split(分隔符)
            
            结果行 = []

            for i in range(索引 - 1, 索引 + 4):
                if 0 <= i < len(行列表):
                    结果行.append(行列表[i])
                else:
                    结果行.append("")  

            return tuple(结果行)
        except ValueError:
            return ("", "", "", "", "") 


#======提取特定行
class 提取特定行:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}),
                "行索引": ("STRING", {"default": "1|2"}),
                "分割字符": ("STRING", {"default": "\n"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING", "STRING", "STRING", "STRING", "STRING", "STRING")
    FUNCTION = "提取特定行"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 提取特定行(self, 输入文本, 行索引, 分割字符):
        if not 分割字符 or 分割字符 == "\n":
            行列表 = 输入文本.split('\n')
        else:
            行列表 = 输入文本.split(分割字符)
        
        索引列表 = [int(索引) - 1 for 索引 in 行索引.split('|') if 索引.isdigit()]
        
        结果列表 = []
        for 索引 in 索引列表:
            if 0 <= 索引 < len(行列表):
                结果列表.append(行列表[索引])
            else:
                结果列表.append("") 
        
        while len(结果列表) < 5:
            结果列表.append("")
        
        非空结果 = [结果 for 结果 in 结果列表[:5] if 结果.strip()]
        if not 分割字符 or 分割字符 == "\n":
            组合结果 = '\n'.join(非空结果)
        else:
            组合结果 = 分割字符.join(非空结果)
        
        return tuple(结果列表[:5] + [组合结果])


#======删除标签内的内容
class 删除字符间内容:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "字符": ("STRING", {"default": "(|)"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "删除字符间内容"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 删除字符间内容(self, 输入文本, 字符):
        try:
            if len(字符) == 3 and 字符[1] == '|':
                开始字符, 结束字符 = 字符[0], 字符[2]
            else:
                return 输入文本  

            模式 = re.escape(开始字符) + '.*?' + re.escape(结束字符)
            结果 = re.sub(模式, '', 输入文本)

            return (结果,)
        except ValueError:
            return ("",)  


#======随机打乱
class 打乱文本行:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "分隔符": ("STRING", {"default": "分隔符"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "打乱文本行"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 打乱文本行(self, 输入文本, 分隔符, any=None):
        if 分隔符 == "":
            行列表 = 输入文本.splitlines()
        elif 分隔符 == "\n":
            行列表 = 输入文本.split("\n")
        else:
            行列表 = 输入文本.split(分隔符)

        行列表 = [行 for 行 in 行列表 if 行.strip()]

        random.shuffle(行列表)

        if 分隔符 == "":
            结果 = "\n".join(行列表)
        elif 分隔符 == "\n":
            结果 = "\n".join(行列表)
        else:
            结果 = 分隔符.join(行列表)

        return (结果,)


#======判断文本包含内容
class 判断文本包含内容:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "原始内容": ("STRING", {"multiline": True, "default": ""}), 
                "检查文本": ("STRING", {"default": "查找字符"}),
                "存在时文本": ("STRING", {"default": "存在返回内容"}),
                "不存在时文本": ("STRING", {"default": "不存在返回内容"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "判断文本包含内容"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 判断文本包含内容(self, 原始内容, 检查文本, 存在时文本, 不存在时文本):
        if not 检查文本:
            return ("",)

        if 检查文本 in 原始内容:
            return (存在时文本,)
        else:
            return (不存在时文本,)


#======文本按条件判断
class 文本条件检查:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "原始内容": ("STRING", {"multiline": True, "default": ""}),  
                "长度条件": ("STRING", {"default": "3-6"}),
                "频率条件": ("STRING", {"default": ""}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("INT", "STRING")
    FUNCTION = "文本条件检查"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 文本条件检查(self, 原始内容, 长度条件, 频率条件):
        长度有效 = self.检查长度条件(原始内容, 长度条件)
        
        频率有效 = self.检查频率条件(原始内容, 频率条件)
        
        if 长度有效 and 频率有效:
            return (1, "1")
        else:
            return (0, "0")

    def 检查长度条件(self, 内容, 条件):
        if '-' in 条件:
            开始, 结束 = map(int, 条件.split('-'))
            return 开始 <= len(内容) <= 结束
        else:
            目标长度 = int(条件)
            return len(内容) == 目标长度

    def 检查频率条件(self, 内容, 条件):
        条件列表 = 条件.split('|')
        for 条件 in 条件列表:
            字符, 计数 = 条件.split(',')
            if 内容.count(字符) != int(计数):
                return False
        return True


#======文本组合
class 文本组合:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "原始文本": ("STRING", {"multiline": True, "default": ""}),
                "组合规则": ("STRING", {"multiline": True, "default": ""}),
                "分割字符": ("STRING", {"default": ""}), 
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING", "STRING", "STRING", "STRING", "STRING")
    FUNCTION = "文本组合"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 文本组合(self, 原始文本, 组合规则, 分割字符):
        if 分割字符:
            原始行 = [行.strip() for 行 in 原始文本.split(分割字符) if 行.strip()]
        else:
            原始行 = [行.strip() for 行 in 原始文本.split('\n') if 行.strip()]

        规则行 = [行.strip() for 行 in 组合规则.split('\n') if 行.strip()]

        输出列表 = []
        for 规则 in 规则行[:5]: 
            结果 = 规则
            for i, 行 in enumerate(原始行, start=1):
                结果 = 结果.replace(f"[{i}]", 行)
            输出列表.append(结果)

        while len(输出列表) < 5:
            输出列表.append("")

        return tuple(输出列表)


#======提取多层指定数据
class 提取特定数据:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}),
                "规则1": ("STRING", {"default": "[3],@|2"}),
                "规则2": ("STRING", {"default": "三,【|】"}),
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "提取特定数据"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 提取特定数据(self, 输入文本, 规则1, 规则2):
        if 规则1.strip():
            return self.按规则1提取(输入文本, 规则1)
        else:
            return self.按规则2提取(输入文本, 规则2)

    def 按规则1提取(self, 输入文本, 规则):
        try:
            行规则, 分割规则 = 规则.split(',')
            分割字符, 组索引 = 分割规则.split('|')
            组索引 = int(组索引) - 1 
        except ValueError:
            return ("",)  

        行列表 = 输入文本.split('\n')
        
        if 行规则.startswith('[') and 行规则.endswith(']'):
            try:
                行索引 = int(行规则[1:-1]) - 1  
                if 0 <= 行索引 < len(行列表):
                    目标行 = 行列表[行索引]
                else:
                    return ("",) 
            except ValueError:
                return ("",)  
        else:
            目标行列表 = [行 for 行 in 行列表 if 行规则 in 行]
            if not 目标行列表:
                return ("",)  
            目标行 = 目标行列表[0]  

        部分 = 目标行.split(分割字符)
        if 0 <= 组索引 < len(部分):
            return (部分[组索引],)
        return ("",)  

    def 按规则2提取(self, 输入文本, 规则):
        try:
            行规则, 标签 = 规则.split(',')
            开始标签, 结束标签 = 标签.split('|')
        except ValueError:
            return ("",)  

        行列表 = 输入文本.split('\n')
        
        if 行规则.startswith('[') and 行规则.endswith(']'):
            try:
                行索引 = int(行规则[1:-1]) - 1  
                if 0 <= 行索引 < len(行列表):
                    目标行 = 行列表[行索引]
                else:
                    return ("",) 
            except ValueError:
                return ("",)  
        else:
            目标行列表 = [行 for 行 in 行列表 if 行规则 in 行]
            if not 目标行列表:
                return ("",)  
            目标行 = 目标行列表[0]  

        开始索引 = 目标行.find(开始标签)
        结束索引 = 目标行.find(结束标签, 开始索引)
        if 开始索引 != -1 and 结束索引 != -1:
            return (目标行[开始索引 + len(开始标签):结束索引],)
        return ("",)  


#======指定字符行参数
class 查找首行内容:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "目标字符": ("STRING", {"default": "数据a"}),  
            },
            "optional": {},
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "查找首行内容"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 查找首行内容(self, 输入文本, 目标字符):
        try:
            行列表 = 输入文本.splitlines()

            for 行 in 行列表:
                if 目标字符 in 行:
                    开始索引 = 行.index(目标字符)
                    结果 = 行[开始索引 + len(目标字符):]
                    return (结果,)

            return ("",)
        except Exception as e:
            return (f"错误: {str(e)}",) 
        

#======获取整数
class 获取整数参数:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": False, "default": "", "forceInput": True}),  
                "目标字符": ("STRING", {"default": ""}),  
            },
            "optional": {},
        }

    RETURN_TYPES = ("INT", "STRING",)
    FUNCTION = "提取整数参数"  # 修正函数名
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 提取整数参数(self, 输入文本, 目标字符):  # 修正方法名
        try:
            行列表 = 输入文本.splitlines()

            for 行 in 行列表:
                if 目标字符 in 行:
                    开始索引 = 行.index(目标字符)
                    结果字符串 = 行[开始索引 + len(目标字符):]
                    try:
                        结果整数 = int(结果字符串)
                    except ValueError:
                        结果整数 = 0  # 返回默认值
                    
                    return (结果整数, 结果字符串)

            return (0, "")  # 返回默认值

        except Exception as e:
            return (0, f"错误: {str(e)}")

#======获取浮点数
class 获取浮点数参数:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": False, "default": "", "forceInput": True}),  
                "目标字符": ("STRING", {"default": ""}),  
            },
            "optional": {},
        }

    RETURN_TYPES = ("FLOAT", "STRING",) 
    FUNCTION = "提取浮点数参数"  # 修正函数名
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 提取浮点数参数(self, 输入文本, 目标字符):  # 修正方法名
        try:
            行列表 = 输入文本.splitlines()

            for 行 in 行列表:
                if 目标字符 in 行:
                    开始索引 = 行.index(目标字符)
                    结果字符串 = 行[开始索引 + len(目标字符):]
                    try:
                        结果浮点数 = float(结果字符串)
                    except ValueError:
                        结果浮点数 = 0.0  # 返回默认值
                    
                    return (结果浮点数, 结果字符串) 

            return (0.0, "")  # 返回默认值

        except Exception as e:
            return (0.0, f"错误: {str(e)}")
        

#======视频指令词模板
class 生成视频提示:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "模式": (["原文本", "文生视频", "图生视频", "首尾帧视频", "视频负面词"],)
            },
            "optional": {}
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "生成提示"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")


    def 生成提示(self, 输入文本, 模式):
        try:
            if 模式 == "原文本":
                return (输入文本,)
                
            elif 模式 == "文生视频":
                前缀 = """您是一位经验丰富的电影导演，擅长创作详细且引人入胜的视觉叙事。在根据用户输入为文生视频生成制作提示时，您的目标是提供精确、按时间顺序排列的描述，以指导生成过程。您的提示应专注于清晰的视觉细节，包括具体的动作、外观、摄像机角度和环境背景。
- 主要动作：以清晰、简洁的场景核心动作或事件描述开始。这应该是视频的焦点。
- 动作和手势：描述场景中的任何动作或手势，无论是来自角色、物体还是环境。包括这些动作如何执行的具体细节。
- 角色或物体的外观：提供任何角色或物体的详细描述，重点关注其身体外观、服装和视觉特征等方面。
- 背景和环境：详细阐述周围环境，突出重要的视觉元素，如景观特征、建筑或重要物体。这些应支持动作并丰富场景。
- 摄像机角度和运动：指定摄像机视角（例如，广角镜头、特写）和任何运动（例如，跟踪、缩放、平移）。
- 灯光和颜色：详细说明灯光设置 - 无论是自然的、人工的还是戏剧性的 - 以及它如何影响场景的氛围。描述有助于营造情绪的色彩色调。
- 突然变化或事件：如果场景中发生任何重大变化（例如，灯光变化、天气变化或情绪变化），请详细描述这些过渡。
通过以这种方式构建您的提示，您可以确保视频输出既引人入胜，又在专业上与用户的预期愿景保持一致。描述应保持在200字限制内，同时保持流畅的流程和电影质量。
以下是我的主要内容：
"""
                return (前缀 + 输入文本,)
                
            elif 模式 == "图生视频":
                前缀 = """您的任务是基于给定图像或用户描述创建电影般、高度详细的视频场景。此提示旨在通过关注精确、按时间顺序的细节来生成沉浸式和视觉动态的视频体验。目标是构建场景的生动和真实描绘，密切关注从主要动作到环境细微差别的每个元素。描述应流畅进行，专注于基本的视觉和电影方面，同时遵守200字限制。
- 主要动作/焦点：
  从场景中核心动作或关键对象的清晰、简洁描述开始。这可能是正在发生的人、物体或事件，提供场景叙事的核心。
- 环境和物体：
  详细描述周围环境或物体。关注它们的纹理、颜色、尺度和位置。这些细节应支持主要动作并有助于场景的氛围。
- 背景细节：
  提供背景的生动描绘。这可能包括自然或建筑元素、远处景观或其他为主角添加上下文的特征。这些细节应丰富视觉叙事。
- 摄像机视角和运动：
  指定使用的摄像机角度或视角 - 无论是广角镜头、特写还是更动态的如跟踪镜头或平移。包括任何摄像机运动，如变焦、倾斜或移动摄影车，如果适用。
- 灯光和颜色：
  详细说明场景中的灯光，解释是自然的、人工的还是两者的结合。考虑灯光如何影响情绪、产生的阴影和色温（暖或冷）。
- 氛围或环境变化：
  如果场景中有任何变化，如天气、灯光或情绪的突然变化，请清楚描述这些过渡。这些环境变化为视频添加了动态元素。
- 最终细节：
  确保所有视觉和上下文元素具有凝聚力并与提供的图像或输入保持一致。确保描述从一点平滑过渡到另一点。
通过遵循此结构，您可以确保场景的每个方面都得到精确处理，提供详细的、电影般的提示，易于转化为视频。保持描述简洁，确保所有视觉和环境因素共同创造流畅且引人入胜的电影体验。
以下是我的主要内容：
"""
                return (前缀 + 输入文本,)
                
            elif 模式 == "首尾帧视频":
                前缀 = """您是一位擅长将静态图像转化为引人入胜的电影序列的专家电影制作人。使用用户提供的两张图像，您的任务是创建一个连接图像一到图像二的无缝视觉叙事。专注于动态过渡，突出按时间顺序展开的动作、环境变化和视觉元素。用电影摄影的语言 crafting 您的描述，确保流畅且沉浸式的叙事。
要求：
 场景连续性：
   - 从图像一设置的详细描述开始，包括中心角色、物体或关键视觉元素。
   - 接着是过渡的流畅叙事，强调图像之间的运动、视觉进展或任何变化。
   - 以图像二关键细节的描述结束，注意环境、角色或视觉构成的演变。
 丰富细节描述：
   - 捕捉角色或主体的显著动作、表情或手势。
   - 描述环境细节，如灯光、色彩调色板、天气和氛围。
   - 融入电影摄影技术，包括摄像机角度、变焦、跟踪镜头或任何动态运动。
 情感和上下文流程：
   - 突出两张图像之间的情感联系或色调转变（例如，从平静到紧张，或从混乱到宁静）。
   - 优先考虑视觉连贯性，即使用户输入与图像之间存在差异。
 输出格式：
   - 从详细描述图像一的核心元素和动作开始。
   - 平滑过渡，描述视觉进展和运动。
   - 以图像二的细节和结论结束。
   - 限制在200字内的单个连贯段落。
以下是我的主要内容：
"""
                return (前缀 + 输入文本,)
                
            elif 模式 == "视频负面词":
                return (
"""过度曝光, 静态伪影, 模糊细节, 可见字幕, 低分辨率绘画, 静态图像, 过于灰色调, 质量差, JPEG压缩伪影, 难看扭曲, 残缺特征, 冗余或额外手指, 渲染不良的手, 描绘不佳的面孔, 解剖畸形, 面部毁容, 畸形肢体, 融合或扭曲的手指, 杂乱和分散注意力的背景元素, 额外或缺失的肢体（例如三条腿）, 过度拥挤的背景与过多人物, 反转或上下颠倒的构图""",)
                
            else:
                return ("",)
                
        except Exception as e:
            return (f"错误: {str(e)}",)