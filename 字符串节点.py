import re
from . import any_typ, note

#====== 1. 按索引提取子字符串
class 按索引提取子字符串:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"default": ""}),
                "索引范围": ("STRING", {"default": "2-6"}),
                "方向": (["从前面", "从后面"], {"default": "从前面"}),
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行提取"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】按字符位置提取。'2-6'表示提取第2到第6个字符。支持从后往前倒数提取。"

    def 执行提取(self, 输入字符串, 索引范围, 方向):
        try:
            if '-' in 索引范围:
                s, e = map(int, 索引范围.split('-'))
            else:
                s = e = int(索引范围)
            s, e = s - 1, e
            if 方向 == "从后面":
                L = len(输入字符串)
                s, e = L - e, L - s
            return (输入字符串[max(0, s):min(len(输入字符串), e)],)
        except: return ("",)

#====== 2. 按分隔符拆分字符串
class 按分隔符拆分字符串:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"default": "文本|内容"}),
                "分隔符": ("STRING", {"default": "|"}),
            }
        }
    RETURN_TYPES = ("STRING", "STRING")
    RETURN_NAMES = ("左侧文本", "右侧文本")
    FUNCTION = "执行拆分"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】根据指定字符将文本一分为二。常用于拆分由'|'或','连接的复合数据。"

    def 执行拆分(self, 输入字符串, 分隔符):
        parts = 输入字符串.split(分隔符, 1)
        return (parts[0], parts[1]) if len(parts) == 2 else (输入字符串, "")

#====== 3. 处理字符串
class 处理字符串:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"multiline": True, "default": ""}),
                "选项": (["不改变", "取数字", "取字母", "转大写", "转小写", "取中文", "去标点", "去换行", "去空行", "去空格", "去格式", "统计字数"], {"default": "不改变"}),
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行处理"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】常用的文本清洗工具。可以快速提取纯数字、过滤掉所有空格换行，或统计当前字数。"

    def 执行处理(self, 输入字符串, 选项):
        if 选项 == "取数字": res = ''.join(re.findall(r'\d', 输入字符串))
        elif 选项 == "取字母": res = ''.join(re.findall(r'[a-zA-Z]', 输入字符串))
        elif 选项 == "转大写": res = 输入字符串.upper()
        elif 选项 == "转小写": res = 输入字符串.lower()
        elif 选项 == "取中文": res = ''.join(re.findall(r'[\u4e00-\u9fff]', 输入字符串))
        elif 选项 == "去标点": res = re.sub(r'[^\w\s\u4e00-\u9fff]', '', 输入字符串)
        elif 选项 == "去换行": res = 输入字符串.replace('\n', '')
        elif 选项 == "去空行": res = '\n'.join([l for l in 输入字符串.splitlines() if l.strip()])
        elif 选项 == "去空格": res = 输入字符串.replace(' ', '')
        elif 选项 == "去格式": res = re.sub(r'\s+', '', 输入字符串)
        elif 选项 == "统计字数": res = str(len(输入字符串))
        else: res = 输入字符串
        return (res,)

#====== 4. 提取前后字符
class 提取前后字符:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"default": ""}),
                "模式": ("STRING", {"default": "标签符"}),
                "位置": (["保留最初之前", "保留最初之后", "保留最后之前", "保留最后之后"], {"default": "保留最初之前"}),
                "包含分隔符": ("BOOLEAN", {"default": False}), 
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行提取"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】以指定字符为锚点进行截取。例如保留'第'字之后的所有内容。"

    def 执行提取(self, 输入字符串, 模式, 位置, 包含分隔符):
        idx = 输入字符串.find(模式) if "最初" in 位置 else 输入字符串.rfind(模式)
        if idx == -1: return ("",)
        if "之前" in 位置:
            return (输入字符串[:idx + len(模式) if 包含分隔符 else idx],)
        else:
            return (输入字符串[idx if 包含分隔符 else idx + len(模式):],)

#====== 5. 简易文本替换
class 简易文本替换:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入字符串": ("STRING", {"multiline": True, "default": "", "forceInput": True}),
                "查找文本": ("STRING", {"default": ""}),
                "替换文本": ("STRING", {"default": ""})
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行替换"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】标准的查找与替换。支持使用'\\n'代表换行符进行替换操作。"

    def 执行替换(self, 输入字符串, 查找文本, 替换文本):
        if not 查找文本: return (输入字符串,)
        rep = 替换文本.replace('\\n', '\n')
        return (输入字符串.replace(查找文本, rep),)

#====== 6. 替换第n次出现
class 替换第n次出现:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "原始文本": ("STRING", {"multiline": True, "default": ""}),
                "出现次数": ("INT", {"default": 1, "min": 0}),
                "查找字符串": ("STRING", {"default": ""}),
                "替换字符串": ("STRING", {"default": ""}),
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行替换"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】精准替换。次数设为0则替换所有，设为2则仅替换第二次出现的字符。"

    def 执行替换(self, 原始文本, 出现次数, 查找字符串, 替换字符串):
        if 出现次数 == 0: return (原始文本.replace(查找字符串, 替换字符串),)
        def repl(m, count=[出现次数]):
            count[0] -= 1
            return 替换字符串 if count[0] == 0 else m.group(0)
        return (re.sub(re.escape(查找字符串), repl, 原始文本),)

#====== 7. 判断是否包含字符
class 判断是否包含字符:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"default": ""}),
                "子字符串": ("STRING", {"default": "查找1|查找2"}),
                "模式": (["同时满足", "任意满足"], {"default": "任意满足"}),
            }
        }
    RETURN_TYPES = ("INT",)
    FUNCTION = "执行检查"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】检查文本中是否存在关键词。'任意满足'类似于逻辑'或'，'同时满足'类似于逻辑'与'。"

    def 执行检查(self, 输入文本, 子字符串, 模式):
        subs = 子字符串.split('|')
        check = all(s in 输入文本 for s in subs) if 模式 == "同时满足" else any(s in 输入文本 for s in subs)
        return (1 if check else 0,)

#====== 8. 段落提取指定索引行
class 段落提取指定索引行:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "行索引": ("STRING", {"default": "1-3"}), 
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行提取"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】按行号提取。'1-3'提取前三行，'1|3|5'提取指定的第1、3、5行。"

    def 执行提取(self, 输入文本, 行索引):
        lines = 输入文本.splitlines()
        res = []
        try:
            if '-' in 行索引:
                s, e = map(int, 行索引.split('-'))
                res = lines[s-1:e]
            else:
                idxs = map(int, 行索引.split('|'))
                res = [lines[i-1] for i in idxs if 0 < i <= len(lines)]
            return ('\n'.join(res),)
        except: return ("",)

#====== 9. 段落提取或移除字符行
class 段落提取或移除字符行:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "子字符串": ("STRING", {"default": "关键词"}), 
                "操作": (["保留", "移除"], {"default": "保留"}), 
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行过滤"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】整行过滤。'保留'包含关键词的行，或'移除'包含关键词的行。"

    def 执行过滤(self, 输入文本, 子字符串, 操作):
        subs = 子字符串.split('|')
        lines = 输入文本.splitlines()
        if 操作 == "保留":
            res = [l for l in lines if any(s in l for s in subs)]
        else:
            res = [l for l in lines if not any(s in l for s in subs)]
        return ('\n'.join(res),)

#====== 10. 按字数过滤行
class 按字数过滤行:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "字数范围": ("STRING", {"default": "5-20"}),  
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行过滤"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】根据每行的长度进行筛选。例如只保留字数在5到20字之间的文本行。"

    def 执行过滤(self, 输入文本, 字数范围):
        try:
            s, e = map(int, 字数范围.split('-'))
            res = [l for l in 输入文本.splitlines() if s <= len(l) <= e]
            return ('\n'.join(res),)
        except: return ("",)

#====== 11. 分割并提取文本
class 分割并提取文本:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}),
                "分隔符": ("STRING", {"default": ","}),
                "索引": ("INT", {"default": 1, "min": 1}),
                "顺序": (["顺序", "倒序"], {"default": "顺序"}),
                "包含分隔符": ("BOOLEAN", {"default": False}), 
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行提取"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】高级拆分工具。将文本按符号切开后，提取其中第N个片段。支持从后往前数。"

    def 执行提取(self, 输入文本, 分隔符, 索引, 顺序, 包含分隔符):
        parts = 输入文本.split(分隔符) if 分隔符 else 输入文本.splitlines()
        if 顺序 == "倒序": parts = parts[::-1]
        if 0 < 索引 <= len(parts):
            res = parts[索引-1]
            if 包含分隔符 and 分隔符:
                if 顺序 == "顺序":
                    if 索引 > 1: res = 分隔符 + res
                    if 索引 < len(parts): res += 分隔符
                else:
                    if 索引 > 1: res += 分隔符
                    if 索引 < len(parts): res = 分隔符 + res
            return (res.strip(),)
        return ("",)

#====== 12. 文本出现次数
class 文本出现次数:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}),
                "字符": ("STRING", {"default": ""}),
            }
        }
    RETURN_TYPES = ("INT", "STRING")
    FUNCTION = "执行统计"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】统计某个词在文中出现了几次。输入'\\n'可统计有效总行数。"

    def 执行统计(self, 输入文本, 字符):
        if 字符 == "\\n":
            count = len([l for l in 输入文本.splitlines() if l.strip()])
        else:
            count = 输入文本.count(字符)
        return (count, str(count))

#====== 13. 删除字符间内容
class 删除字符间内容:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "符号对": ("STRING", {"default": "(|)"}),
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行删除"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】删除括号内或标签内的内容。'符号对'格式：'(|)'表示删除所有括号及其内部文本。"

    def 执行删除(self, 输入文本, 符号对):
        if '|' not in 符号对: return (输入文本,)
        s, e = 符号对.split('|', 1)
        pattern = re.escape(s) + '.*?' + re.escape(e)
        return (re.sub(pattern, '', 输入文本),)

#====== 14. 判断文本包含内容
class 判断文本包含内容:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "原始内容": ("STRING", {"multiline": True, "default": ""}), 
                "检查文本": ("STRING", {"default": ""}),
                "存在时文本": ("STRING", {"default": "True"}),
                "不存在时文本": ("STRING", {"default": "False"}),
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行判断"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】条件返回。如果文中包含关键词，则输出A文本，否则输出B文本。"

    def 执行判断(self, 原始内容, 检查文本, 存在时文本, 不存在时文本):
        return (存在时文本 if 检查文本 in 原始内容 else 不存在时文本,)

#====== 15. 文本条件检查
class 文本条件检查:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "原始内容": ("STRING", {"multiline": True, "default": ""}),  
                "长度条件": ("STRING", {"default": "1-100"}),
                "频率条件": ("STRING", {"default": ""}),
            }
        }
    RETURN_TYPES = ("INT", "STRING")
    FUNCTION = "执行检查"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】双重校验。长度条件如'10-50'；频率条件如'的,2'（表示字符'的'必须出现2次）。"

    def 执行检查(self, 原始内容, 长度条件, 频率条件):
        ok_len = False
        if '-' in 长度条件:
            s, e = map(int, 长度条件.split('-'))
            ok_len = s <= len(原始内容) <= e
        else: ok_len = len(原始内容) == int(长度条件)
        
        ok_freq = True
        if 频率条件:
            for item in 频率条件.split('|'):
                if ',' in item:
                    char, count = item.split(',')
                    if 原始内容.count(char) != int(count): ok_freq = False
        return (1 if ok_len and ok_freq else 0, "1" if ok_len and ok_freq else "0")

#====== 16. 提取特定数据 (多层级)
class 提取特定数据:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}),
                "规则1": ("STRING", {"default": "[1],:|2"}),
                "规则2": ("STRING", {"default": ""}),
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行提取"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】从复杂文本行提取。规则示例'[2],:|1'指：在第2行，按冒号分割，取第1个片段。"

    def 执行提取(self, 输入文本, 规则1, 规则2):
        rule = 规则1 if 规则1.strip() else 规则2
        if not rule: return ("",)
        try:
            line_idx, split_part = rule.split(',')
            line_idx = int(line_idx[1:-1]) - 1
            lines = 输入文本.splitlines()
            if line_idx >= len(lines): return ("",)
            target = lines[line_idx]
            
            if '|' in split_part:
                sep, idx = split_part.split('|')
                return (target.split(sep)[int(idx)-1],)
            return (target,)
        except: return ("",)

#====== 17. 查找首行内容
class 查找首行内容:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"multiline": True, "default": ""}), 
                "目标字符": ("STRING", {"default": ""}),  
            }
        }
    RETURN_TYPES = ("STRING",)
    FUNCTION = "执行查找"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】在段落中搜索关键词，并返回该行中关键词之后的所有内容。"

    def 执行查找(self, 输入文本, 目标字符):
        for line in 输入文本.splitlines():
            if 目标字符 in line:
                return (line.split(目标字符, 1)[1],)
        return ("",)

#====== 18. 获取整数参数
class 获取整数参数:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"forceInput": True}),  
                "目标字符": ("STRING", {"default": ""}),  
            }
        }
    RETURN_TYPES = ("INT", "STRING")
    FUNCTION = "提取整数"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】从文本行中提取数字。例如输入'高度: 512'，目标设为'高度:'，将返回整数512。"

    def 提取整数(self, 输入文本, 目标字符):
        for line in 输入文本.splitlines():
            if 目标字符 in line:
                val_str = line.split(目标字符, 1)[1].strip()
                try: return (int(re.search(r'-?\d+', val_str).group()), val_str)
                except: pass
        return (0, "")

#====== 19. 获取浮点数参数
class 获取浮点数参数:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "输入文本": ("STRING", {"forceInput": True}),  
                "目标字符": ("STRING", {"default": ""}),  
            }
        }
    RETURN_TYPES = ("FLOAT", "STRING")
    FUNCTION = "提取浮点数"
    CATEGORY = "【Excel】联动插件/字符串节点"
    DESCRIPTION = "【使用方法】从文本行中提取小数。常用于从导出的参数文本中提取权重或缩放值。"

    def 提取浮点数(self, 输入文本, 目标字符):
        for line in 输入文本.splitlines():
            if 目标字符 in line:
                val_str = line.split(目标字符, 1)[1].strip()
                try: return (float(re.search(r'-?\d+\.?\d*', val_str).group()), val_str)
                except: pass
        return (0.0, "")
