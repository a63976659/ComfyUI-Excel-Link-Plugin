import os, re, io, base64, csv, torch, shutil, requests, chardet, pathlib
import openpyxl, folder_paths, node_helpers
import numpy as np
from PIL import Image, ImageOps, ImageSequence, ImageDraw, ImageFont
from pathlib import Path
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage
from io import BytesIO
from . import any_typ, note

# 定义常量
COMFYUI_OUTPUT_DIR = "output"

#======全能加载图像
class 全能加载图像:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "图像输入": ("STRING", {"default": ""}),
            }
        }
    RETURN_TYPES = ("IMAGE", "MASK")
    FUNCTION = "加载图像"
    OUTPUT_NODE = False
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 加载图像(self, 图像输入):
        路径 = 图像输入.strip()
        for 控制符 in ['\u202a', '\u202b', '\u202c', '\u202d', '\u202e']:
            while 路径.startswith(控制符):
                路径 = 路径.lstrip(控制符)

        来源 = None
        小写路径 = 路径.lower()
        if 小写路径.startswith('http://') or 小写路径.startswith('https://'):
            来源 = '网络'
        elif os.path.isfile(路径):
            来源 = '本地'
        else:
            来源 = 'base64'
            
        if 来源 == '本地':
            图像 = Image.open(路径)
        elif 来源 == '网络':
            响应 = requests.get(路径)
            响应.raise_for_status()
            图像 = Image.open(io.BytesIO(响应.content))
        else:  # 'base64'
            if ',' in 路径 and 路径.startswith('data:'):
                _, 数据 = 路径.split(',', 1)
            else:
                数据 = 路径
            解码数据 = base64.b64decode(数据)
            图像 = Image.open(io.BytesIO(解码数据))

        图像 = 图像.convert('RGBA')
        有透明通道 = 图像.mode == 'RGBA'
        if 有透明通道:
            透明通道 = np.array(图像)[:, :, 3]
            遮罩 = (透明通道 > 0).astype(np.float32)
            遮罩张量 = torch.from_numpy(遮罩).unsqueeze(0).unsqueeze(0) 
        else:
            遮罩张量 = torch.zeros((1, 1, 图像.size[1], 图像.size[0]), dtype=torch.float32) 
            
        np图像 = np.array(图像).astype(np.float32) / 255.0
        图像张量 = torch.from_numpy(np图像).unsqueeze(0)

        return 图像张量, 遮罩张量


#======加载重置图像
class 加载重置图像:
    @classmethod
    def INPUT_TYPES(s):
        输入目录 = folder_paths.get_input_directory()
        文件列表 = [文件.name for 文件 in Path(输入目录).iterdir() if 文件.is_file()]
        return {
            "required": {
                "图像": (sorted(文件列表), {"image_upload": True}),
                "最大尺寸": ("INT", {"default": 0, "min": 0, "max": 4096, "step": 8}),
                "尺寸选项": (["保持不变", "自定义", "百万像素", "小", "中", "大", 
                                "480P-横屏(视频 4:3)", "480P-竖屏(视频 3:4)", "720P-横屏(视频 16:9)", "720P-竖屏(视频 9:16)", "832×480", "480×832"], 
                                {"default": "保持不变"})
            }
        }

    RETURN_TYPES = ("IMAGE", "MASK", "STRING")
    RETURN_NAMES = ("图像", "遮罩", "信息")
    FUNCTION = "加载图像"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 加载图像(self, 图像, 最大尺寸, 尺寸选项):
        图像路径 = folder_paths.get_annotated_filepath(图像)
        图像对象 = Image.open(图像路径)
        宽度, 高度 = 图像对象.size
        宽高比 = 宽度 / 高度

        def 获取目标尺寸():
            if 尺寸选项 == "保持不变":
                return 宽度, 高度
            elif 尺寸选项 == "百万像素":
                return self._调整到百万像素(宽度, 高度)
            elif 尺寸选项 == "自定义":
                比例 = min(最大尺寸 / 宽度, 最大尺寸 / 高度)
                return round(宽度 * 比例), round(高度 * 比例)
            
            尺寸选项字典 = {
                "小": (
                    (768, 512) if 宽高比 >= 1.23 else
                    (512, 768) if 宽高比 <= 0.82 else
                    (768, 768)
                ),
                "中": (
                    (1216, 832) if 宽高比 >= 1.23 else
                    (832, 1216) if 宽高比 <= 0.82 else
                    (1216, 1216)
                ),
                "大": (
                    (1600, 1120) if 宽高比 >= 1.23 else
                    (1120, 1600) if 宽高比 <= 0.82 else
                    (1600, 1600)
                ),
                "百万像素": self._调整到百万像素(宽度, 高度),
                "480P-横屏(视频 4:3)": (640, 480),
                "480P-竖屏(视频 3:4)": (480, 640),
                "720P-横屏(视频 16:9)": (1280, 720),
                "720P-竖屏(视频 9:16)": (720, 1280),
                "832×480": (832, 480),
                "480×832": (480, 832),
            }
            return 尺寸选项字典[尺寸选项]
        
        目标宽度, 目标高度 = 获取目标尺寸()
        输出图像列表 = []
        输出遮罩列表 = []

        for 帧 in ImageSequence.Iterator(图像对象):
            帧 = ImageOps.exif_transpose(帧)
            if 帧.mode == 'P':
                帧 = 帧.convert("RGBA")
            elif 'A' in 帧.getbands():
                帧 = 帧.convert("RGBA")
            
            if 尺寸选项 == "保持不变":
                图像帧 = 帧.convert("RGB")
            else:
                if 尺寸选项 == "自定义" or 尺寸选项 == "百万像素":
                    比例 = min(目标宽度 / 宽度, 目标高度 / 高度)
                    调整宽度 = round(宽度 * 比例)
                    调整高度 = round(高度 * 比例)
                    图像帧 = 帧.convert("RGB").resize((调整宽度, 调整高度), Image.Resampling.BILINEAR)
                else:
                    图像帧 = 帧.convert("RGB")
                    图像帧 = ImageOps.fit(图像帧, (目标宽度, 目标高度), method=Image.Resampling.BILINEAR, centering=(0.5, 0.5))

            图像数组 = np.array(图像帧).astype(np.float32) / 255.0
            输出图像列表.append(torch.from_numpy(图像数组)[None,])

            # 处理遮罩
            if 'A' in 帧.getbands():
                遮罩帧 = 帧.getchannel('A')
                if 尺寸选项 == "自定义" or 尺寸选项 == "百万像素":
                    遮罩帧 = 遮罩帧.resize((调整宽度, 调整高度), Image.Resampling.BILINEAR)
                else:
                    遮罩帧 = ImageOps.fit(遮罩帧, (目标宽度, 目标高度), method=Image.Resampling.BILINEAR, centering=(0.5, 0.5))
                遮罩 = np.array(遮罩帧).astype(np.float32) / 255.0
                遮罩 = 1. - 遮罩
            else:
                if 尺寸选项 == "自定义" or 尺寸选项 == "百万像素":
                    遮罩 = np.zeros((调整高度, 调整宽度), dtype=np.float32)
                else:
                    遮罩 = np.zeros((目标高度, 目标宽度), dtype=np.float32)
            输出遮罩列表.append(torch.from_numpy(遮罩).unsqueeze(0))
        
        输出图像 = torch.cat(输出图像列表, dim=0) if len(输出图像列表) > 1 else 输出图像列表[0]
        输出遮罩 = torch.cat(输出遮罩列表, dim=0) if len(输出遮罩列表) > 1 else 输出遮罩列表[0]
        信息 = f"图像路径: {图像路径}\n原始尺寸: {宽度}x{高度}\n调整后尺寸: {目标宽度}x{目标高度}"
        return (输出图像, 输出遮罩, 信息)

    @classmethod
    def VALIDATE_INPUTS(s, 图像):
        if not folder_paths.exists_annotated_filepath(图像):
            return f"无效的图像文件: {图像}"
        return True
        
    def _调整到百万像素(self, 宽度, 高度):
        宽高比 = 宽度 / 高度
        目标面积 = 1000000  # 1百万像素
        if 宽高比 > 1:  # 横屏
            新宽度 = int(np.sqrt(目标面积 * 宽高比))
            新高度 = int(目标面积 / 新宽度)
        else:  # 竖屏
            新高度 = int(np.sqrt(目标面积 / 宽高比))
            新宽度 = int(目标面积 / 新高度)
            
        新宽度 = (新宽度 + 7) // 8 * 8
        新高度 = (新高度 + 7) // 8 * 8
        return 新宽度, 新高度


#======重置图像
class 重置图像:
    @classmethod
    def INPUT_TYPES(s):
        return {
            "required": {
                "图像": ("IMAGE",),
                "遮罩": ("MASK",),
                "最大尺寸": ("INT", {"default": 0, "min": 0, "max": 4096, "step": 8}),
                "尺寸选项": ([
                    "自定义", "百万像素", "小", "中", "大", 
                    "480P-横屏(视频 4:3)", "480P-竖屏(视频 3:4)", "720P-横屏(视频 16:9)", "720P-竖屏(视频 9:16)", "832×480", "480×832"], 
                    {"default": "百万像素"}
                )
            }
        }

    RETURN_TYPES = ("IMAGE", "MASK", "INT", "INT")
    RETURN_NAMES = ("图像", "遮罩", "宽度", "高度")
    FUNCTION = "处理图像"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 处理图像(self, 图像, 遮罩, 最大尺寸=1024, 尺寸选项="自定义"):
        批次大小 = 图像.shape[0]
        处理后的图像列表 = []
        处理后的遮罩列表 = []
        宽度列表 = []
        高度列表 = []

        for i in range(批次大小):
            当前图像 = 图像[i]
            当前遮罩 = 遮罩[i]
            
            输入图像对象 = Image.fromarray((当前图像.numpy() * 255).astype(np.uint8))
            输入遮罩对象 = Image.fromarray((当前遮罩.numpy() * 255).astype(np.uint8))
            
            宽度, 高度 = 输入图像对象.size

            处理后的图像对象 = 输入图像对象.copy()
            处理后的图像对象 = ImageOps.exif_transpose(处理后的图像对象)

            处理后的遮罩对象 = 输入遮罩对象.copy()
            处理后的遮罩对象 = ImageOps.exif_transpose(处理后的遮罩对象)

            if 处理后的图像对象.mode == 'P':
                处理后的图像对象 = 处理后的图像对象.convert("RGBA")
            elif 'A' in 处理后的图像对象.getbands():
                处理后的图像对象 = 处理后的图像对象.convert("RGBA")

            if 处理后的遮罩对象.mode != "L":
                处理后的遮罩对象 = 处理后的遮罩对象.convert("L")

            if 尺寸选项 != "自定义":
                宽高比 = 宽度 / 高度

                尺寸选项字典 = {
                    "小": (
                        (768, 512) if 宽高比 >= 1.23 else
                        (512, 768) if 宽高比 <= 0.82 else
                        (768, 768)
                    ),
                    "中": (
                        (1216, 832) if 宽高比 >= 1.23 else
                        (832, 1216) if 宽高比 <= 0.82 else
                        (1216, 1216)
                    ),
                    "大": (
                        (1600, 1120) if 宽高比 >= 1.23 else
                        (1120, 1600) if 宽高比 <= 0.82 else
                        (1600, 1600)
                    ),
                    "百万像素": self._调整到百万像素(宽度, 高度),
                    "480P-横屏(视频 4:3)": (640, 480),
                    "480P-竖屏(视频 3:4)": (480, 640),
                    "720P-横屏(视频 16:9)": (1280, 720),
                    "720P-竖屏(视频 9:16)": (720, 1280),
                    "832×480": (832, 480),
                    "480×832": (480, 832)
                }

                目标宽度, 目标高度 = 尺寸选项字典[尺寸选项]
                处理后的图像对象 = 处理后的图像对象.convert("RGB")
                处理后的图像对象 = ImageOps.fit(处理后的图像对象, (目标宽度, 目标高度), method=Image.Resampling.BILINEAR, centering=(0.5, 0.5))
                
                处理后的遮罩对象 = ImageOps.fit(处理后的遮罩对象, (目标宽度, 目标高度), method=Image.Resampling.BILINEAR, centering=(0.5, 0.5))
            else:
                比例 = min(最大尺寸 / 宽度, 最大尺寸 / 高度)
                调整宽度 = round(宽度 * 比例)
                调整高度 = round(高度 * 比例)

                处理后的图像对象 = 处理后的图像对象.convert("RGB")
                处理后的图像对象 = 处理后的图像对象.resize((调整宽度, 调整高度), Image.Resampling.BILINEAR)
                
                处理后的遮罩对象 = 处理后的遮罩对象.resize((调整宽度, 调整高度), Image.Resampling.BILINEAR)

            处理后的图像数组 = np.array(处理后的图像对象).astype(np.float32) / 255.0
            处理后的图像张量 = torch.from_numpy(处理后的图像数组)
            处理后的图像列表.append(处理后的图像张量)

            处理后的遮罩数组 = np.array(处理后的遮罩对象).astype(np.float32) / 255.0
            处理后的遮罩张量 = torch.from_numpy(处理后的遮罩数组)
            处理后的遮罩列表.append(处理后的遮罩张量)

            if 尺寸选项 != "自定义":
                宽度列表.append(目标宽度)
                高度列表.append(目标高度)
            else:
                宽度列表.append(调整宽度)
                高度列表.append(调整高度)

        输出图像 = torch.stack(处理后的图像列表)
        输出遮罩 = torch.stack(处理后的遮罩列表)
        
        if all(宽 == 宽度列表[0] for 宽 in 宽度列表) and all(高 == 高度列表[0] for 高 in 高度列表):
            return (输出图像, 输出遮罩, 宽度列表[0], 高度列表[0])
        else:
            return (输出图像, 输出遮罩, 宽度列表[0], 高度列表[0])

    def _调整到百万像素(self, 宽度, 高度):
        宽高比 = 宽度 / 高度
        目标面积 = 1000000
        
        if 宽高比 > 1:
            新宽度 = int(np.sqrt(目标面积 * 宽高比))
            新高度 = int(目标面积 / 新宽度)
        else:
            新高度 = int(np.sqrt(目标面积 / 宽高比))
            新宽度 = int(目标面积 / 新高度)

        新宽度 = (新宽度 + 7) // 8 * 8
        新高度 = (新高度 + 7) // 8 * 8
        
        return 新宽度, 新高度


#======裁剪图像
class 裁剪图像:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "图像": ("IMAGE",),
                "遮罩": ("MASK",),
                "宽度": ("INT", {"default": 768, "min": 0, "max": 4096, "step": 8}),
                "高度": ("INT", {"default": 768, "min": 0, "max": 4096, "step": 8}),
            }
        }

    RETURN_TYPES = ("IMAGE", "MASK", "INT", "INT")
    RETURN_NAMES = ("图像", "遮罩", "宽度", "高度")
    FUNCTION = "处理图像"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 处理图像(self, 图像, 遮罩, 宽度=768, 高度=768):
        输入图像 = Image.fromarray((图像.squeeze(0).numpy() * 255).astype(np.uint8))
        输入遮罩 = Image.fromarray((遮罩.squeeze(0).numpy() * 255).astype(np.uint8))
        
        原始宽度, 原始高度 = 输入图像.size

        处理后的图像列表 = []
        处理后的遮罩列表 = []

        for 帧 in [输入图像]:
            帧 = ImageOps.exif_transpose(帧)

            if 帧.mode == 'P':
                帧 = 帧.convert("RGBA")
            elif 'A' in 帧.getbands():
                帧 = 帧.convert("RGBA")

            处理后的图像 = 帧.convert("RGB")
            处理后的图像 = ImageOps.fit(处理后的图像, (宽度, 高度), method=Image.Resampling.BILINEAR, centering=(0.5, 0.5))

            处理后的图像数组 = np.array(处理后的图像).astype(np.float32) / 255.0
            处理后的图像张量 = torch.from_numpy(处理后的图像数组)[None,]
            处理后的图像列表.append(处理后的图像张量)

        # 处理遮罩
        输入遮罩 = ImageOps.exif_transpose(输入遮罩)
        处理后的遮罩 = 输入遮罩.convert("L")
        处理后的遮罩 = ImageOps.fit(处理后的遮罩, (宽度, 高度), method=Image.Resampling.BILINEAR, centering=(0.5, 0.5))
        处理后的遮罩数组 = np.array(处理后的遮罩).astype(np.float32) / 255.0
        处理后的遮罩张量 = torch.from_numpy(处理后的遮罩数组)[None,]
        处理后的遮罩列表.append(处理后的遮罩张量)

        输出图像 = torch.cat(处理后的图像列表, dim=0) if len(处理后的图像列表) > 1 else 处理后的图像列表[0]
        输出遮罩 = torch.cat(处理后的遮罩列表, dim=0) if len(处理后的遮罩列表) > 1 else 处理后的遮罩列表[0]

        return (输出图像, 输出遮罩, 宽度, 高度)


#======保存图像
class 保存图像:
    def __init__(self):
        self.输出目录 = folder_paths.get_output_directory()

    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "图像": ("IMAGE",),
                "保存路径": ("STRING", {"default": "./output"}),
                "图像名称": ("STRING", {"default": "ComfyUI"}),
                "图像格式": (["PNG", "JPG"], {"default": "JPG"})
            },
        }

    RETURN_TYPES = ("IMAGE",)
    RETURN_NAMES = ("图像",)
    FUNCTION = "保存图像"
    OUTPUT_NODE = True
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 保存图像(self, 图像, 保存路径, 图像名称, 图像格式):
        if not isinstance(图像, torch.Tensor):
            raise ValueError("无效的图像张量格式")
        if 保存路径 == "./output":
            保存路径 = self.输出目录
        elif not os.path.isabs(保存路径):
            保存路径 = os.path.join(self.输出目录, 保存路径)
        os.makedirs(保存路径, exist_ok=True)
        
        # 移除可能存在的扩展名
        基本名称 = os.path.splitext(图像名称)[0]
        
        批次大小 = 图像.shape[0]
        通道到模式 = {1: "L", 3: "RGB", 4: "RGBA"}

        for i in range(批次大小):
            if 图像格式 == "PNG":
                文件名 = f"{基本名称}.png" if 批次大小 == 1 else f"{基本名称}_{i:05d}.png"
                保存格式 = "PNG"
                保存参数 = {"compress_level": 0}
            else:  # JPG
                文件名 = f"{基本名称}.jpg" if 批次大小 == 1 else f"{基本名称}_{i:05d}.jpg"
                保存格式 = "JPEG"
                保存参数 = {"quality": 100}
            
            完整路径 = os.path.join(保存路径, 文件名)
            单张图像 = 图像[i].cpu().numpy()
            单张图像 = (单张图像 * 255).astype('uint8')
            通道数 = 单张图像.shape[2]
            if 通道数 not in 通道到模式:
                raise ValueError(f"不支持的通道数: {通道数}")
            模式 = 通道到模式[通道数]
            if 通道数 == 1:
                单张图像 = 单张图像[:, :, 0]
            图像对象 = Image.fromarray(单张图像, 模式)
            
            if 图像格式 == "JPG":
                图像对象 = 图像对象.convert("RGB")
            
            图像对象.save(完整路径, format=保存格式, **保存参数)
        return (图像,)


#======文件操作
class 文件操作:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "源文件路径": ("STRING", {"default": "", "multiline": False}),
                "目标文件路径": ("STRING", {"default": "", "multiline": False}),
                "操作类型": (["复制", "剪切"], {"default": "复制"}),
            },
            "optional": {"任意": (any_typ,)}
        }

    RETURN_TYPES = ("STRING",)
    RETURN_NAMES = ("结果",)
    FUNCTION = "复制剪切文件"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 复制剪切文件(self, 源文件路径, 目标文件路径, 操作类型, any=None):
        结果 = "执行失败"
        try:
            if not os.path.isfile(源文件路径):
                raise FileNotFoundError(f"源文件未找到: {源文件路径}")

            os.makedirs(os.path.dirname(目标文件路径), exist_ok=True)

            if 操作类型 == "复制":
                shutil.copy2(源文件路径, 目标文件路径)
                结果 = "执行成功: 文件已复制"
            elif 操作类型 == "剪切":
                shutil.move(源文件路径, 目标文件路径)
                结果 = "执行成功: 文件已剪切"
            else:
                raise ValueError("操作类型无效，仅支持 '复制' 或 '剪切'。")
        except FileNotFoundError as e:
            结果 = f"执行失败: {str(e)}"
        except ValueError as e:
            结果 = f"执行失败: {str(e)}"
        except Exception as e:
            结果 = f"执行失败: {str(e)}"

        return (结果,)


#======替换文件名
class 替换文件名:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "文件路径": ("STRING", {"default": "path/to/your/file.jpg"}),
                "新文件名": ("STRING", {"default": ""}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "替换文件名"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 替换文件名(self, 文件路径, 新文件名, any=None):
        目录名 = os.path.dirname(文件路径)
        _, 文件扩展名 = os.path.splitext(文件路径)
        新文件名 = self.清理文件名(新文件名)
        新文件路径 = os.path.join(目录名, 新文件名 + 文件扩展名)
        return (新文件路径,)
        
    def 清理文件名(self, 文件名):
        无效字符 = r'[\/:*?"<>|]'
        return re.sub(无效字符, '_', 文件名)


# 将类名改为与 __init__.py 匹配
class 写入文本文件:  # 原来是 文本写入文件
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "文本内容": ("STRING", {"default": "", "multiline": True}),
                "文件路径": ("STRING", {"default": "path/to/your/file.txt"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "写入文本"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 写入文本(self, 文本内容, 文件路径, any=None):
        try:
            目录路径 = os.path.dirname(文件路径)
            if not os.path.exists(目录路径):
                os.makedirs(目录路径)
            文件存在 = os.path.exists(文件路径)
            模式 = 'a' if 文件存在 else 'w'
            
            with open(文件路径, 模式, encoding='utf-8') as 文件:
                if 文件存在:
                    文件.write('\n\n') 
                文件.write(文本内容)
            return ("写入成功: " + 文本内容,)
        except Exception as e:
            return (f"错误: {str(e)}",)

#======清理文件
class 清理文件:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "删除项目": ("STRING", {"default": "33.png\ncs1/01.png\ncs1", "multiline": True}),
            },
            "optional": {"任意": (any_typ,)} 
        }
    
    RETURN_TYPES = ("STRING",)
    RETURN_NAMES = ("结果",)
    FUNCTION = "删除文件"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 删除文件(self, 删除项目, any=None):
        结果 = "执行成功: 所有指定项已从output目录删除"
        错误信息列表 = []
        基础输出目录 = Path.cwd() / COMFYUI_OUTPUT_DIR
        项目列表 = 删除项目.strip().split('\n')

        for 项目 in 项目列表:
            项目 = 项目.strip()
            if not 项目:
                continue
            if 项目 == "[DeleteAll]":
                try:
                    for 文件或目录 in 基础输出目录.glob('*'):
                        if 文件或目录.is_file() or 文件或目录.is_symlink():
                            文件或目录.unlink()
                        elif 文件或目录.is_dir():
                            shutil.rmtree(文件或目录)
                    continue
                except Exception as e:
                    错误信息列表.append(f"从output目录删除全部失败: {str(e)}")
                    continue
            目标路径 = 基础输出目录 / 项目
            try:
                目标路径.relative_to(基础输出目录)
            except ValueError:
                错误信息列表.append(f"{项目} 不在output目录范围内，无法删除")
                continue
            if not 目标路径.exists():
                错误信息列表.append(f"在output目录下找不到 {项目}")
                continue
            try:
                if 目标路径.is_file() or 目标路径.is_symlink():
                    目标路径.unlink()
                elif 目标路径.is_dir():
                    shutil.rmtree(目标路径)
            except Exception as e:
                错误信息列表.append(f"从output目录删除 {项目} 失败: {str(e)}")
        if 错误信息列表:
            结果 = "部分执行失败:\n" + "\n".join(错误信息列表)
        return (结果,)


#======文件路径和后缀统计
class 文件路径和后缀统计:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "文件夹路径": ("STRING",),
                "文件扩展名": (["jpg", "png", "jpg&png", "txt", "csv", "全部"], {"default": "jpg"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING", "INT", "LIST")
    FUNCTION = "文件列表和后缀统计"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 文件列表和后缀统计(self, 文件夹路径, 文件扩展名, any=None):
        try:
            if not os.path.isdir(文件夹路径):
                return ("", 0, [])

            if 文件扩展名 == "全部":
                文件路径列表 = [os.path.join(文件夹路径, 文件) for 文件 in os.listdir(文件夹路径) if os.path.isfile(os.path.join(文件夹路径, 文件))]
            elif 文件扩展名 == "jpg&png":
                文件路径列表 = [os.path.join(文件夹路径, 文件) for 文件 in os.listdir(文件夹路径) if 文件.lower().endswith(('.jpg', '.png'))]
            else:
                文件路径列表 = [os.path.join(文件夹路径, 文件) for 文件 in os.listdir(文件夹路径) if 文件.lower().endswith('.' + 文件扩展名)]

            return ('\n'.join(文件路径列表), len(文件路径列表), 文件路径列表)
        except Exception as e:
            return ("", 0, [])


#======文字图像
class 文字图像:
    @classmethod
    def INPUT_TYPES(cls):
        字体目录 = os.path.join(os.path.dirname(os.path.realpath(__file__)), "fonts")
        if not os.path.exists(字体目录):
            os.makedirs(字体目录)
            字体文件列表 = []
        else:
            字体文件列表 = [文件 for 文件 in os.listdir(字体目录) if 文件.lower().endswith(('.ttf', '.otf'))]
        字体文件列表 = 字体文件列表 or ["arial.ttf"]
        return {
            "required": {
                "文本": ("STRING", {"default": "Hello, World!", "multiline": True}),
                "字体": (字体文件列表, ),
                "最大宽度": ("INT", {"default": 300, "min": 1, "max": 2048, "step": 1}),
                "字体属性": ("STRING", {"default": "#FFFFFF,1,1,C,1", "multiline": False}),
                "字体描边": ("STRING", {"default": "#000000,2,1", "multiline": False}),
                "字体背景": ("STRING", {"default": "#333333,5,10,1", "multiline": False})
            },
            "optional": {"任意": (any_typ,)} 
        }
    RETURN_TYPES = ("IMAGE",)
    FUNCTION = "生成文字图像"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 生成文字图像(self, 文本, 字体, 最大宽度, 字体属性, 字体描边, 字体背景):
        字体目录 = os.path.join(os.path.dirname(os.path.realpath(__file__)), "fonts")
        字体路径 = os.path.join(字体目录, 字体)
        try:
            字体对象 = ImageFont.truetype(字体路径, 1)
        except Exception as e:
            return None

        绘图对象 = ImageDraw.Draw(Image.new('RGBA', (1, 1)))
        行列表 = 文本.split("\n")
        最大文本宽度 = 0
        文本高度 = 0
        for 行 in 行列表:
            左, 上, 右, 下 = 绘图对象.textbbox((0, 0), 行, font=字体对象)
            行宽度 = 右 - 左
            行高度 = 下 - 上
            最大文本宽度 = max(最大文本宽度, 行宽度)
            文本高度 += 行高度

        if 最大文本宽度 == 0:
            最大文本宽度 = 1

        比例 = 最大宽度 / 最大文本宽度
        新字体大小 = int(1 * 比例)
        字体对象 = ImageFont.truetype(字体路径, 新字体大小)

        字体颜色 = "#FFFFFF"
        字母间距因子 = 1.0
        行间距因子 = 1.0
        对齐方式 = "C"
        不透明度 = 1.0
        描边颜色 = "#000000"
        描边宽度 = 0.0
        描边不透明度 = 1.0
        背景颜色 = "#333333"
        扩展宽度 = 5
        圆角半径 = 10
        背景不透明度 = 0.9

        try:
            if 字体属性.strip() == "":
                字体颜色 = "#FFFFFF"
                字母间距因子 = 1.0
                行间距因子 = 1.0
                对齐方式 = "C"
                不透明度 = 1.0
            else:
                属性列表 = 字体属性.split(',')
                if len(属性列表) >= 5:
                    字体颜色 = 属性列表[0].strip()
                    字母间距因子 = float(属性列表[1]) if 属性列表[1] else 1.0
                    行间距因子 = float(属性列表[2]) if 属性列表[2] else 1.0
                    对齐方式 = 属性列表[3].strip().upper()
                    不透明度 = float(属性列表[4]) if 属性列表[4] else 1.0
                else:
                    字体颜色 = "#FFFFFF"
                    字母间距因子 = 1.0
                    行间距因子 = 1.0
                    对齐方式 = "C"
                    不透明度 = 1.0
        except:
            字体颜色 = "#FFFFFF"
            字母间距因子 = 1.0
            行间距因子 = 1.0
            对齐方式 = "C"
            不透明度 = 1.0

        try:
            if 字体描边.strip() == "":
                描边宽度 = 0.0
            else:
                描边属性列表 = 字体描边.split(',')
                if len(描边属性列表) >= 3:
                    描边颜色 = 描边属性列表[0].strip()
                    描边宽度 = float(描边属性列表[1]) if 描边属性列表[1] else 1.0
                    描边不透明度 = float(描边属性列表[2]) if 描边属性列表[2] else 1.0
                else:
                    描边宽度 = 0.0
        except:
            描边宽度 = 0.0

        try:
            if 字体背景.strip() == "":
                背景颜色 = None
            else:
                背景属性列表 = 字体背景.split(',')
                if len(背景属性列表) >= 4:
                    背景颜色 = 背景属性列表[0].strip()
                    扩展宽度 = int(背景属性列表[1]) if 背景属性列表[1] else 5
                    圆角半径 = int(背景属性列表[2]) if 背景属性列表[2] else 10
                    背景不透明度 = float(背景属性列表[3]) if 背景属性列表[3] else 0.9
                else:
                    背景颜色 = None
        except:
            背景颜色 = None

        实际最大宽度 = 0
        for 行 in 行列表:
            行宽度 = 0
            for 字符 in 行:
                字符宽度 = 绘图对象.textbbox((0, 0), 字符, font=字体对象)[2]
                行宽度 += 字符宽度 + (字体对象.size * (字母间距因子 - 1))
            实际最大宽度 = max(实际最大宽度, 行宽度)

        if 实际最大宽度 > 最大宽度:
            比例 = 最大宽度 / 实际最大宽度
            新字体大小 = int(新字体大小 * 比例)
            字体对象 = ImageFont.truetype(字体路径, 新字体大小)

        字体上升, 字体下降 = 字体对象.getmetrics()
        行高度 = 字体上升 + 字体下降

        if len(行列表) > 1:
            文本高度 = 行高度 * (len(行列表) - 1) * 行间距因子 + 行高度
        else:
            文本高度 = 行高度

        图像宽度 = 最大宽度
        图像高度 = int(文本高度 + 新字体大小 * 0.2)
        图像 = Image.new('RGBA', (图像宽度, 图像高度), (0, 0, 0, 0))
        绘图对象 = ImageDraw.Draw(图像)

        if 背景颜色 is not None:
            try:
                背景颜色元组 = (
                    int(背景颜色[1:3], 16),
                    int(背景颜色[3:5], 16),
                    int(背景颜色[5:7], 16),
                    int(背景不透明度 * 255)
                )
                绘图对象.rounded_rectangle(
                    [0, 0, 图像宽度, 图像高度],
                    fill=背景颜色元组,
                    radius=圆角半径
                )
            except:
                pass

        文本Y坐标 = 新字体大小 * 0.1
        try:
            字体颜色元组 = (
                int(字体颜色[1:3], 16),
                int(字体颜色[3:5], 16),
                int(字体颜色[5:7], 16),
                int(不透明度 * 255)
            )
            描边颜色元组 = (
                int(描边颜色[1:3], 16),
                int(描边颜色[3:5], 16),
                int(描边颜色[5:7], 16),
                int(描边不透明度 * 255)
            )
        except:
            字体颜色元组 = (255, 255, 255, 255)
            描边颜色元组 = (0, 0, 0, 255)

        for i, 行 in enumerate(行列表):
            行宽度 = 0
            for 字符 in 行:
                字符宽度 = 绘图对象.textbbox((0, 0), 字符, font=字体对象)[2]
                行宽度 += 字符宽度 + (字体对象.size * (字母间距因子 - 1))
            行宽度 = max(行宽度, 1)

            if 对齐方式 == "L":
                x = 0
            elif 对齐方式 == "R":
                x = 最大宽度 - 行宽度
            else:
                x = (最大宽度 - 行宽度) / 2

            if 描边宽度 > 0:
                for sx in range(-int(描边宽度), int(描边宽度) + 1):
                    for sy in range(-int(描边宽度), int(描边宽度) + 1):
                        if sx == 0 and sy == 0:
                            continue
                        字符x = x + sx
                        字符y = 文本Y坐标 + sy
                        for 字符 in 行:
                            字符宽度 = 绘图对象.textbbox((0, 0), 字符, font=字体对象)[2]
                            绘图对象.text((字符x, 字符y), 字符, font=字体对象, fill=描边颜色元组)
                            字符x += 字符宽度 + (字体对象.size * (字母间距因子 - 1))

            字符x = x
            for 字符 in 行:
                字符宽度 = 绘图对象.textbbox((0, 0), 字符, font=字体对象)[2]
                绘图对象.text((字符x, 文本Y坐标), 字符, font=字体对象, fill=字体颜色元组)
                字符x += 字符宽度 + (字体对象.size * (字母间距因子 - 1))

            if i < len(行列表) - 1:
                文本Y坐标 += 行高度 * 行间距因子
            else:
                文本Y坐标 += 行高度

        图像数据 = np.array(图像)
        透明通道 = 图像数据[:, :, 3]
        非零索引 = np.where(透明通道 > 0)
        if len(非零索引[0]) > 0:
            最小Y = np.min(非零索引[0])
            最大Y = np.max(非零索引[0])
            最小X = np.min(非零索引[1])
            最大X = np.max(非零索引[1])
            图像 = 图像.crop((最小X, 最小Y, 最大X + 1, 最大Y + 1))
        else:
            pass

        文本内容宽度 = 最大X - 最小X + 1 if len(非零索引[0]) > 0 else 最大宽度

        图像宽度, 图像高度 = 图像.size
        if 文本内容宽度 < 最大宽度:
            新图像 = Image.new('RGBA', (最大宽度, 图像高度), (0, 0, 0, 0))
            新绘图对象 = ImageDraw.Draw(新图像)
            x偏移 = (最大宽度 - 文本内容宽度) // 2
            新图像.paste(图像, (x偏移, 0))
            图像 = 新图像

        图像宽度, 图像高度 = 图像.size
        if 图像宽度 > 最大宽度:
            高度比例 = 图像高度 / 图像宽度
            图像 = 图像.resize((最大宽度, int(最大宽度 * 高度比例)))

        图像数组 = np.array(图像).astype(np.float32) / 255.0
        图像张量 = torch.from_numpy(图像数组).unsqueeze(0)

        return (图像张量,)


#======图像层叠加
class 图像层叠加:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "图像1": ("IMAGE", {"forceInput": True}),
                "图像2": ("IMAGE", {"forceInput": True}),
                "对齐方式": (["左上", "中上", "右上", "左下", "中下", "右下", "居中"], ),
                "缩放比例": ("FLOAT", {"default": 1.0, "min": 0.1, "max": 10.0, "step": 0.1}),
                "不透明度": ("FLOAT", {"default": 1.0, "min": 0.0, "max": 1.0, "step": 0.1}),
                "偏移": ("STRING", {"default": "0,0,0,0", "multiline": False})
            }
        }

    RETURN_TYPES = ("IMAGE",)
    FUNCTION = "叠加图像"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 叠加图像(self, 图像1, 图像2, 对齐方式, 偏移, 缩放比例, 不透明度):
        图像1数组 = 图像1.cpu().numpy().squeeze()
        图像2数组 = 图像2.cpu().numpy().squeeze()
        图像1对象 = Image.fromarray((图像1数组 * 255).astype(np.uint8)).convert('RGBA')
        图像2对象 = Image.fromarray((图像2数组 * 255).astype(np.uint8)).convert('RGBA')
        图像1宽度, 图像1高度 = 图像1对象.size
        新宽度 = int(图像1宽度 * 缩放比例)
        新高度 = int(图像1高度 * 缩放比例)
        图像1对象 = 图像1对象.resize((新宽度, 新高度), Image.LANCZOS)
        图像1对象 = self.调整不透明度(图像1对象, 不透明度)
        图像1宽度, 图像1高度 = 图像1对象.size
        图像2宽度, 图像2高度 = 图像2对象.size
        最大宽度 = max(图像1宽度, 图像2宽度)
        最大高度 = max(图像1高度, 图像2高度)
        画布 = Image.new('RGBA', (最大宽度, 最大高度), (0, 0, 0, 0))
        if 对齐方式 == "左上":
            图像2x, 图像2y = 0, 0
        elif 对齐方式 == "中上":
            图像2x = (最大宽度 - 图像2宽度) // 2
            图像2y = 0
        elif 对齐方式 == "右上":
            图像2x = 最大宽度 - 图像2宽度
            图像2y = 0
        elif 对齐方式 == "左下":
            图像2x = 0
            图像2y = 最大高度 - 图像2高度
        elif 对齐方式 == "中下":
            图像2x = (最大宽度 - 图像2宽度) // 2
            图像2y = 最大高度 - 图像2高度
        elif 对齐方式 == "右下":
            图像2x = 最大宽度 - 图像2宽度
            图像2y = 最大高度 - 图像2高度
        elif 对齐方式 == "居中":
            图像2x = (最大宽度 - 图像2宽度) // 2
            图像2y = (最大高度 - 图像2高度) // 2

        右偏移, 左偏移, 下偏移, 上偏移 = 0, 0, 0, 0
        偏移列表 = 偏移.split(',')
        if len(偏移列表) >= 4:
            try:
                右偏移 = int(偏移列表[0]) if 偏移列表[0] else 0
                左偏移 = int(偏移列表[1]) if 偏移列表[1] else 0
                下偏移 = int(偏移列表[2]) if 偏移列表[2] else 0
                上偏移 = int(偏移列表[3]) if 偏移列表[3] else 0
            except ValueError:
                pass
        if 对齐方式 == "左上":
            图像1x, 图像1y = 0, 0
        elif 对齐方式 == "中上":
            图像1x = (最大宽度 - 图像1宽度) // 2
            图像1y = 0
        elif 对齐方式 == "右上":
            图像1x = 最大宽度 - 图像1宽度
            图像1y = 0
        elif 对齐方式 == "左下":
            图像1x = 0
            图像1y = 最大高度 - 图像1高度
        elif 对齐方式 == "中下":
            图像1x = (最大宽度 - 图像1宽度) // 2
            图像1y = 最大高度 - 图像1高度
        elif 对齐方式 == "右下":
            图像1x = 最大宽度 - 图像1宽度
            图像1y = 最大高度 - 图像1高度
        elif 对齐方式 == "居中":
            图像1x = (最大宽度 - 图像1宽度) // 2
            图像1y = (最大高度 - 图像1高度) // 2
            图像1x += 右偏移 - 左偏移
            图像1y += 下偏移 - 上偏移
            图像1x = max(0, min(图像1x, 最大宽度 - 图像1宽度))
            图像1y = max(0, min(图像1y, 最大高度 - 图像1高度))
            图像2x = max(0, min(图像2x, 最大宽度 - 图像2宽度))
            图像2y = max(0, min(图像2y, 最大高度 - 图像2高度))
        画布.paste(图像2对象, (图像2x, 图像2y), 图像2对象.split()[-1])
        画布.paste(图像1对象, (图像1x, 图像1y), 图像1对象.split()[-1])
        输出图像数组 = np.array(画布).astype(np.float32) / 255.0
        输出张量 = torch.from_numpy(输出图像数组).unsqueeze(0)
        return (输出张量,)

    def 调整不透明度(self, 图像对象, 不透明度):
        if 不透明度 < 1.0:
            图像对象 = 图像对象.copy() 
            透明度 = np.array(图像对象.split()[-1]) * 不透明度
            透明度 = 透明度.astype(np.uint8)
            图像对象.putalpha(Image.fromarray(透明度))
        return 图像对象


#======读取表格数据
class 读取Excel数据:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "表格路径": ("STRING", {"default": "F:\ComfyUI与Excel联动方案：示例文件\武侠：开局奖励满级神功第三集.xlsx"}),
                "工作表名称": ("STRING", {"default": "Sheet1"}),
                "行范围": ("STRING", {"default": "2-3"}),
                "列范围": ("STRING", {"default": "1-4"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "读取表格数据"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 读取表格数据(self, 表格路径, 工作表名称, 行范围, 列范围, any=None):
        try:
            if "-" in 行范围:
                开始行, 结束行 = map(int, 行范围.split("-"))
            else:
                开始行 = 结束行 = int(行范围)

            if "-" in 列范围:
                开始列, 结束列 = map(int, 列范围.split("-"))
            else:
                开始列 = 结束列 = int(列范围)

            开始行 = max(1, 开始行)
            开始列 = max(1, 开始列)

            工作簿 = openpyxl.load_workbook(表格路径, read_only=True, data_only=True)
            工作表 = 工作簿[工作表名称]

            输出行列表 = []
            for 行 in range(开始行, 结束行 + 1):
                行数据 = []
                for 列 in range(开始列, 结束列 + 1):
                    单元格值 = 工作表.cell(row=行, column=列).value
                    行数据.append(str(单元格值) if 单元格值 is not None else "")
                输出行列表.append("|".join(行数据))

            工作簿.close()
            del 工作簿

            return ("\n".join(输出行列表),)

        except Exception as e:
            return (f"错误: {str(e)}",)


#======写入表格数据
class 写入Excel数据:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "表格路径": ("STRING", {"default": "F:\ComfyUI与Excel联动方案：示例文件\武侠：开局奖励满级神功第三集.xlsx"}),
                "工作表名称": ("STRING", {"default": "Sheet1"}), 
                "行范围": ("STRING", {"default": "2-3"}),
                "列范围": ("STRING", {"default": "1-4"}),
                "数据": ("STRING", {"default": "", "multiline": True}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "写入表格数据"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 写入表格数据(self, 表格路径, 工作表名称, 行范围, 列范围, 数据, any=None):
        try:
            if not os.path.exists(表格路径):
                return (f"错误: 文件不存在: {表格路径}",)

            if not os.access(表格路径, os.W_OK):
                return (f"错误: 没有写入权限: {表格路径}",)

            if "-" in 行范围:
                开始行, 结束行 = map(int, 行范围.split("-"))
            else:
                开始行 = 结束行 = int(行范围)

            if "-" in 列范围:
                开始列, 结束列 = map(int, 列范围.split("-"))
            else:
                开始列 = 结束列 = int(列范围)

            开始行 = max(1, 开始行)
            开始列 = max(1, 开始列)

            工作簿 = openpyxl.load_workbook(表格路径, read_only=False, data_only=True)
            工作表 = 工作簿[工作表名称]

            数据行列表 = 数据.strip().split("\n")

            for 行索引, 行数据 in enumerate(数据行列表, start=开始行):
                if 行索引 > 结束行:
                    break

                单元格值列表 = 行数据.split("|")
                for 列索引, 单元格值 in enumerate(单元格值列表, start=开始列):
                    if 列索引 > 结束列:
                        break

                    if 单元格值.strip():
                        工作表.cell(row=行索引, column=列索引).value = 单元格值.strip()

            工作簿.save(表格路径)

            工作簿.close()
            del 工作簿

            return ("数据写入成功!",)

        except PermissionError as pe:
            return (f"权限错误: {str(pe)}",)
        except Exception as e:
            return (f"错误: {str(e)}",)


#======图片插入表格
class 写入Excel图片:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "表格路径": ("STRING", {"default": "F:\ComfyUI与Excel联动方案：示例文件\武侠：开局奖励满级神功第三集.xlsx"}),
                "工作表名称": ("STRING", {"default": "Sheet1"}),
                "行范围": ("STRING", {"default": "1"}),
                "列范围": ("STRING", {"default": "1"}),
                "图片路径": ("STRING", {"default": "path/to/your/image.png"}),
            }
        }

    RETURN_TYPES = ("STRING",)
    FUNCTION = "写入Excel图片"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    
    def IS_CHANGED(self, 表格路径, 工作表名称, 行范围, 列范围, 图片路径):
        """
        修复 IS_CHANGED 方法 - 参数必须与 INPUT_TYPES 中定义的完全一致
        """
        return float("NaN")

    def 写入Excel图片(self, 表格路径, 工作表名称, 行范围, 列范围, 图片路径):
        try:
            # 处理可能的列表输入 - 提取第一个值
            def 提取单个值(输入值):
                if isinstance(输入值, list):
                    if len(输入值) > 0:
                        return str(输入值[0])
                    else:
                        return ""
                else:
                    return str(输入值)
            
            # 提取单个值
            表格路径 = 提取单个值(表格路径)
            工作表名称 = 提取单个值(工作表名称)
            行范围 = 提取单个值(行范围)
            列范围 = 提取单个值(列范围)
            图片路径 = 提取单个值(图片路径)
            
            print(f"调试信息: 表格路径 = {表格路径}")
            print(f"调试信息: 图片路径 = {图片路径}")
            
            # 现在进行规范化路径
            表格路径 = os.path.normpath(表格路径)
            图片路径 = os.path.normpath(图片路径)
            
            print(f"调试信息: 规范化后表格路径 = {表格路径}")
            print(f"调试信息: 规范化后图片路径 = {图片路径}")
            
            # 检查基础权限
            if not os.path.exists(表格路径):
                return (f"错误: Excel文件不存在: {表格路径}",)
                
            # 检查目录权限
            目录 = os.path.dirname(表格路径)
            print(f"调试信息: 目录 = {目录}")
            
            if not os.path.exists(目录):
                return (f"错误: 目录不存在: {目录}",)
                
            # 检查目录写入权限
            if not os.access(目录, os.W_OK):
                return (f"错误: 没有目录写入权限: {目录}",)

            # 检查文件是否被占用
            try:
                with open(表格路径, 'rb') as f:
                    pass
            except PermissionError:
                return ("错误: Excel文件被其他进程锁定。请关闭Excel和其他可能使用此文件的应用程序。",)
            except Exception as e:
                print(f"调试信息: 文件访问检查异常 - {str(e)}")

            # 关键检查：确保图片路径是文件而不是目录
            if os.path.isdir(图片路径):
                return (f"错误: 图片路径指向的是文件夹，不是图片文件: {图片路径}",)
                
            if not os.path.exists(图片路径):
                return (f"错误: 图片文件不存在: {图片路径}",)
                
            if not os.access(图片路径, os.R_OK):
                return (f"错误: 没有图片文件读取权限: {图片路径}",)
                
            # 检查文件扩展名
            有效扩展名 = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff'}
            文件扩展名 = os.path.splitext(图片路径)[1].lower()
            if 文件扩展名 not in 有效扩展名:
                return (f"错误: 不支持的文件格式: {文件扩展名}。支持的格式: {', '.join(有效扩展名)}",)

            # 解析行列范围
            def 解析范围(范围字符串):
                if isinstance(范围字符串, list):
                    范围字符串 = 范围字符串[0] if 范围字符串 else "1"
                
                范围字符串 = str(范围字符串).strip()
                if "-" in 范围字符串:
                    开始, 结束 = map(int, 范围字符串.split("-"))
                    return 开始, 结束
                else:
                    值 = int(范围字符串)
                    return 值, 值
            
            开始行, 结束行 = 解析范围(行范围)
            开始列, 结束列 = 解析范围(列范围)
                
            开始行 = max(1, 开始行)
            开始列 = max(1, 开始列)

            print(f"调试信息: 开始加载工作簿...")
            
            # 加载工作簿
            工作簿 = openpyxl.load_workbook(表格路径, read_only=False, data_only=True)
            
            # 确保工作表存在
            if 工作表名称 not in 工作簿.sheetnames:
                return (f"错误: 工作表 '{工作表名称}' 不存在。可用工作表: {', '.join(工作簿.sheetnames)}",)
                
            工作表 = 工作簿[工作表名称]
            
            print(f"调试信息: 工作簿加载成功，开始处理图片...")
            
            # 处理图片
            缩略图尺寸 = (128, 128)
            with PILImage.open(图片路径) as 图片对象:
                图片对象 = 图片对象.resize(缩略图尺寸)
                图片字节数组 = BytesIO()
                图片对象.save(图片字节数组, format='PNG')
                图片字节数组.seek(0)
                openpyxl图片对象 = OpenpyxlImage(图片字节数组)

            # 插入图片
            单元格地址 = get_column_letter(开始列) + str(开始行)
            工作表.add_image(openpyxl图片对象, 单元格地址)
            
            print(f"调试信息: 图片插入完成，开始保存...")
            
            # 方案A: 直接保存
            try:
                工作簿.save(表格路径)
                工作簿.close()
                print(f"调试信息: 直接保存成功")
                return ("图片插入成功!",)
            except PermissionError:
                print(f"调试信息: 直接保存失败，尝试临时文件方案...")
                工作簿.close()
                
                # 方案B: 使用临时文件方案
                import tempfile
                import shutil
                
                临时目录 = tempfile.gettempdir()
                临时文件 = os.path.join(临时目录, f"temp_excel_{os.getpid()}_{os.path.basename(表格路径)}")
                
                print(f"调试信息: 临时文件路径 = {临时文件}")
                
                # 重新加载工作簿到临时文件
                工作簿2 = openpyxl.load_workbook(表格路径, read_only=False, data_only=True)
                工作表2 = 工作簿2[工作表名称]
                
                # 重新插入图片
                工作表2.add_image(openpyxl图片对象, 单元格地址)
                
                工作簿2.save(临时文件)
                工作簿2.close()
                
                # 复制回原位置
                shutil.copy2(临时文件, 表格路径)
                os.remove(临时文件)
                
                print(f"调试信息: 临时文件方案成功")
                return ("图片插入成功!",)
                
        except PermissionError as pe:
            return (f"权限错误: {str(pe)}",)
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            print(f"详细错误信息: {error_details}")
            return (f"错误: {str(e)}",)


#======查找表格数据
class 查找Excel数据:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "表格路径": ("STRING", {"default": "F:\ComfyUI与Excel联动方案：示例文件\武侠：开局奖励满级神功第三集.xlsx"}),
                "工作表名称": ("STRING", {"default": "Sheet1"}),
                "查找内容": ("STRING", {"default": "查找内容"}),
                "查找模式": (["精确查找", "模糊查找"], {"default": "精确查找"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("STRING", "INT", "INT")
    FUNCTION = "查找表格数据"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 查找表格数据(self, 表格路径, 工作表名称, 查找内容, 查找模式, any=None):
        try:
            if not os.path.exists(表格路径):
                return (f"错误: 文件不存在: {表格路径}", 0, 0)
            if not os.access(表格路径, os.R_OK):
                return (f"错误: 没有读取权限: {表格路径}",0, 0)
                
            工作簿 = openpyxl.load_workbook(表格路径, read_only=True, data_only=True)
            工作表 = 工作簿[工作表名称]

            结果列表 = []
            找到的行 = 0
            找到的列 = 0
            for 行 in range(1, 工作表.max_row + 1):
                for 列 in range(1, 工作表.max_column + 1):
                    单元格 = 工作表.cell(row=行, column=列)
                    单元格值 = 单元格.value if 单元格.value is not None else ""
                    单元格值字符串 = str(单元格值)
                    if (查找模式 == "精确查找" and 单元格值字符串 == 查找内容) or \
                       (查找模式 == "模糊查找" and 查找内容 in 单元格值字符串):
                        结果列表.append(f"{工作表名称}|{行}|{列}|{单元格值}")
                        找到的行 = 行
                        找到的列 = 列

            工作簿.close()
            del 工作簿
            if not 结果列表:
                return ("未找到结果。", 0, 0)
            return ("\n".join(结果列表), 找到的行, 找到的列)
        except Exception as e:
            return (f"错误: {str(e)}", 0, 0)


#======读取表格数量差
class 读取Excel行列差:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "表格路径": ("STRING", {"default": "F:\ComfyUI与Excel联动方案：示例文件\武侠：开局奖励满级神功第三集.xlsx"}),
                "工作表名称": ("STRING", {"default": "Sheet1"}),
                "读取模式": (["读行", "读列"], {"default": "读行"}),
                "索引": ("STRING", {"default": "1,3"}),
            },
            "optional": {"任意": (any_typ,)} 
        }

    RETURN_TYPES = ("INT",)
    FUNCTION = "读取表格行或列数量差"
    CATEGORY = "【Excel】联动插件/文件处理节点"
    DESCRIPTION = note
    
    def IS_CHANGED(self, **kwargs): 
        return float("NaN")

    def 读取表格行或列数量差(self, 表格路径, 工作表名称, 读取模式, 索引, any=None):
        try:
            if not os.path.exists(表格路径):
                return (0,)  # 修正：返回默认整数值

            if not os.access(表格路径, os.R_OK):
                return (0,)  # 修正：返回默认整数值

            工作簿 = openpyxl.load_workbook(表格路径, read_only=True, data_only=True)
            工作表 = 工作簿[工作表名称]

            def 统计单元格(模式, 索引值):
                数量 = 0
                if 模式 == "读行":
                    for 列 in range(1, 工作表.max_column + 1):
                        单元格值 = 工作表.cell(row=索引值, column=列).value
                        if 单元格值 is not None:
                            数量 += 1
                        else:
                            break
                elif 模式 == "读列":
                    for 行 in range(1, 工作表.max_row + 1):
                        单元格值 = 工作表.cell(row=行, column=索引值).value
                        if 单元格值 is not None:
                            数量 += 1
                        else:
                            break
                return 数量

            索引 = 索引.strip()
            if "," in 索引:
                try:
                    索引1, 索引2 = map(int, 索引.split(","))
                except ValueError:
                    return (0,)  # 修正：返回默认整数值

                数量1 = 统计单元格(读取模式, 索引1)
                数量2 = 统计单元格(读取模式, 索引2)
                结果 = 数量1 - 数量2
            else:
                try:
                    索引值 = int(索引)
                except ValueError:
                    return (0,)  # 修正：返回默认整数值

                结果 = 统计单元格(读取模式, 索引值)

            工作簿.close()
            del 工作簿

            return (结果,)

        except Exception as e:
            return (0,)  # 修正：返回默认整数值