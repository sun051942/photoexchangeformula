
# -*- coding: utf-8 -*-
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from PIL import Image, ImageTk, ImageOps
import pyperclip
from docx import Document
from formula_to_clipboard import copy_image_to_clipboard
from latex_cleaner import extract_main_formula
from latex_to_omml import insert_latex_to_word
import threading

# 检测是否在打包的.exe环境中运行
def is_frozen():
    """检测是否在PyInstaller打包的环境中"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

# 获取应用程序路径
def get_app_path():
    """Return the folder containing the executable or source script."""
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    else:
        return os.path.dirname(os.path.abspath(__file__))

# pix2tex 默认配置内容
PIX2TEX_CONFIG = """gpu_devices: null
backbone_layers:
- 2
- 3
- 7
betas:
- 0.9
- 0.999
batchsize: 64
bos_token: 1
channels: 1
data: dataset/data/train.pkl
debug: false
decoder_args:
  attn_on_attn: true
  cross_attend: true
  ff_glu: true
  rel_pos_bias: false
  use_scalenorm: false
dim: 256
encoder_depth: 4
eos_token: 2
epochs: 10
gamma: 0.9995
heads: 8
id: null
load_chkpt: null
lr: 0.001
lr_step: 30
max_height: 192
max_seq_len: 512
max_width: 672
micro_batchsize: -1
min_height: 32
min_width: 32
model_path: checkpoints
name: pix2tex
num_layers: 4
num_tokens: 8000
optimizer: Adam
output_path: outputs
pad: false
pad_token: 0
patch_size: 16
sample_freq: 3000
save_freq: 5
scheduler: StepLR
seed: 42
encoder_structure: hybrid
temperature: 0.2
test_samples: 5
testbatchsize: 20
tokenizer: dataset/tokenizer.json
valbatches: 100
valdata: dataset/data/val.pkl
"""

# 设置模型缓存目录
def setup_model_cache():
    """Keep runtime caches beside the portable application."""
    cache_base = os.path.join(get_app_path(), "cache", "pix2tex")
    cache_dir = os.path.join(cache_base, "models")

    # 创建所有必要的目录
    try:
        os.makedirs(cache_dir, exist_ok=True)
        os.makedirs(os.path.join(cache_base, "huggingface"), exist_ok=True)
        os.makedirs(os.path.join(cache_base, "transformers"), exist_ok=True)

        # 创建settings目录和config.yaml文件（pix2tex需要）
        settings_dir = os.path.join(cache_base, "settings")
        os.makedirs(settings_dir, exist_ok=True)

        # 创建model/settings目录
        model_settings_dir = os.path.join(cache_base, "model", "settings")
        os.makedirs(model_settings_dir, exist_ok=True)

        # 写入配置文件到两个位置
        for config_path in [
            os.path.join(settings_dir, "config.yaml"),
            os.path.join(model_settings_dir, "config.yaml")
        ]:
            if not os.path.exists(config_path):
                with open(config_path, 'w', encoding='utf-8') as f:
                    f.write(PIX2TEX_CONFIG)

        print(f"模型缓存目录: {cache_dir}")
    except Exception as e:
        print(f"创建缓存目录失败: {e}")

    # 设置环境变量
    os.environ['TORCH_HOME'] = cache_base
    os.environ['HF_HOME'] = os.path.join(cache_base, "huggingface")
    os.environ['XDG_CACHE_HOME'] = cache_base
    os.environ['HUGGINGFACE_HUB_CACHE'] = os.path.join(cache_base, "huggingface", "hub")

    # 修改当前工作目录到缓存目录，让pix2tex能找到settings/config.yaml
    try:
        os.chdir(cache_base)
        print(f"工作目录已切换到: {cache_base}")
    except Exception as e:
        print(f"切换工作目录失败: {e}")

    return cache_dir

# 初始化模型缓存
MODEL_CACHE_DIR = setup_model_cache()

# 设置中文字体支持
try:
    # 尝试导入字体模块
    from tkinter.font import Font
except ImportError:
    pass

# 设置控制台编码为UTF-8，避免中文乱码
if sys.platform == 'win32':
    try:
        # 检查stdout和stderr是否为None（窗口模式下可能为None）
        if hasattr(sys, 'stdout') and sys.stdout is not None:
            if hasattr(sys.stdout, 'reconfigure'):
                sys.stdout.reconfigure(encoding='utf-8')
        if hasattr(sys, 'stderr') and sys.stderr is not None:
            if hasattr(sys.stderr, 'reconfigure'):
                sys.stderr.reconfigure(encoding='utf-8')
    except:
        # Python 3.6 及以下版本或窗口模式下的处理
        try:
            import codecs
            if hasattr(sys, 'stdout') and sys.stdout is not None and hasattr(sys.stdout, 'buffer'):
                sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
            if hasattr(sys, 'stderr') and sys.stderr is not None and hasattr(sys.stderr, 'buffer'):
                sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')
        except:
            # 在窗口模式下忽略编码设置错误
            pass

try:
    from pix2tex.cli import LatexOCR
except ImportError as exc:
    raise RuntimeError("请先安装依赖：python -m pip install -r requirements.txt") from exc


class FormulaRecognizerGUI:
    """图形界面的公式识别器"""

    def __init__(self, root):
        self.root = root
        self.root.title("数学公式识别工具")
        self.root.geometry("850x650")
        self.root.minsize(800, 600)

        # 设置主题颜色
        self.colors = {
            'primary': '#4A90E2',
            'secondary': '#50C878',
            'accent': '#FF6B6B',
            'background': '#F5F7FA',
            'foreground': '#333333',
            'light_bg': '#FFFFFF',
            'border': '#E0E0E0',
            'success': '#50C878',
            'warning': '#FFA500',
            'error': '#FF6B6B',
            'text_light': '#666666',
            'shadow': '#00000015'
        }

        # 设置全局背景色
        self.root.configure(bg=self.colors['background'])

        # 创建加载动画标志
        self.loading = False
        self.loading_label = None

        # 添加应用程序图标（使用Unicode字符作为图标）
        try:
            # 尝试设置窗口图标（如果有合适的.ico文件）
            # self.root.iconbitmap("app_icon.ico")
            pass
        except:
            # 如果失败，继续运行
            pass

        # 添加全局样式配置
        self.setup_global_styles()

        self.model = None
        self.current_image_path = None
        self.current_latex = None
        self.crop_box = None

        self.setup_ui()

    @staticmethod
    def clean_latex(latex_str):
        """仅去掉成对的外层数学定界符，保留公式的全部结构。"""
        return extract_main_formula(latex_str)

    def setup_global_styles(self):
        """设置全局样式"""
        # 配置ttk样式
        self.style = ttk.Style()

        # 配置按钮样式
        self.style.configure(
            "Modern.TButton",
            padding=(12, 6),
            font=("微软雅黑", 10),
            foreground=self.colors['foreground'],
            background=self.colors['primary'],
            borderwidth=1
        )

        # 配置按钮悬停效果
        self.style.map(
            "Modern.TButton",
            foreground=[('pressed', self.colors['light_bg']),
                       ('active', self.colors['primary'])],
            background=[('pressed', self.colors['primary']),
                       ('active', self.colors['border'])]
        )

        # 配置成功按钮样式
        self.style.configure(
            "Success.TButton",
            padding=(12, 6),
            font=("微软雅黑", 10),
            foreground=self.colors['foreground'],
            background=self.colors['success']
        )

        self.style.map(
            "Success.TButton",
            foreground=[('disabled', self.colors['text_light']),
                        ('pressed', self.colors['foreground']),
                        ('active', self.colors['foreground'])],
            background=[('pressed', self.colors['secondary']),
                       ('active', self.colors['secondary'])]
        )

    def setup_ui(self):
        """设置用户界面"""
        # 创建主容器
        main_frame = tk.Frame(self.root, bg=self.colors['background'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # 标题区域 - 添加图标和更现代的标题设计
        title_frame = tk.Frame(main_frame, bg=self.colors['background'])
        title_frame.pack(fill=tk.X, pady=(0, 20))

        # 标题图标（使用Unicode数学符号）
        icon_label = tk.Label(
            title_frame,
            text="∑",  # 求和符号作为简单图标
            font=("Arial", 28),
            bg=self.colors['background'],
            fg=self.colors['primary']
        )
        icon_label.pack(side=tk.LEFT, padx=(0, 10))

        # 标题
        title_label = tk.Label(
            title_frame,
            text="数学公式识别工具",
            font=("微软雅黑", 18, "bold"),
            bg=self.colors['background'],
            fg=self.colors['primary']
        )
        title_label.pack(side=tk.LEFT, anchor=tk.CENTER)

        # 功能介绍
        intro_frame = tk.Frame(main_frame, bg=self.colors['background'])
        intro_frame.pack(fill=tk.X, pady=(0, 15))

        intro_text = "带坐标轴或水印的图片请先框选公式，再识别；结果可修改并导出为可编辑Word公式。"
        intro_label = tk.Label(
            intro_frame,
            text=intro_text,
            font=("微软雅黑", 10),
            bg=self.colors['background'],
            fg=self.colors['text_light'],
            wraplength=750
        )
        intro_label.pack(anchor=tk.W)

        # 按钮框架 - 添加阴影效果
        button_frame = tk.Frame(main_frame, bg=self.colors['background'])
        button_frame.pack(pady=(0, 20), fill=tk.X)

        # 选择图片按钮
        self.select_btn = ttk.Button(
            button_frame,
            text="选择图片",
            command=self.select_image,
            style="Modern.TButton",
            width=12
        )
        self.select_btn.pack(side=tk.LEFT, padx=5)

        self.crop_btn = ttk.Button(
            button_frame,
            text="框选公式区域",
            command=self.select_formula_region,
            style="Modern.TButton",
            state=tk.DISABLED,
            width=14
        )
        self.crop_btn.pack(side=tk.LEFT, padx=5)

        # 识别按钮
        self.recognize_btn = ttk.Button(
            button_frame,
            text="识别公式",
            command=self.recognize_formula,
            style="Modern.TButton",
            state=tk.DISABLED,
            width=12
        )
        self.recognize_btn.pack(side=tk.LEFT, padx=5)

        # 复制LaTeX按钮
        self.copy_latex_btn = ttk.Button(
            button_frame,
            text="复制LaTeX",
            command=self.copy_latex,
            style="Modern.TButton",
            state=tk.DISABLED,
            width=12
        )
        self.copy_latex_btn.pack(side=tk.LEFT, padx=5)

        # 复制图片按钮（使用不同颜色区分）
        self.style.configure(
            "Success.TButton",
            padding=(12, 6),
            font=("微软雅黑", 10)
        )

        self.copy_img_btn = ttk.Button(
            button_frame,
            text="复制图片到剪贴板",
            command=self.copy_image,
            style="Success.TButton",
            state=tk.DISABLED,
            width=18
        )
        self.copy_img_btn.pack(side=tk.LEFT, padx=5)

        self.export_btn = ttk.Button(
            main_frame,
            text="保存可编辑Word公式",
            command=self.export_word_formula,
            style="Modern.TButton",
            state=tk.DISABLED,
        )
        self.export_btn.pack(anchor=tk.W, pady=(0, 8))

        # 内容区域框架
        content_frame = tk.Frame(main_frame, bg=self.colors['background'])
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 图片显示区域 - 添加阴影和圆角效果
        self.image_frame = tk.LabelFrame(
            content_frame,
            text="图片预览",
            font=("微软雅黑", 10),
            padx=15,
            pady=15,
            bg=self.colors['light_bg'],
            relief=tk.FLAT,
            bd=1,
            highlightbackground=self.colors['border'],
            highlightthickness=1
        )
        # 模拟圆角效果和阴影（Tkinter原生不支持真正的圆角）
        self.image_frame.pack(side=tk.LEFT, padx=(0, 10), fill=tk.BOTH, expand=True)

        # 添加简单的装饰线条
        decor_line = tk.Frame(
            self.image_frame,
            height=3,
            bg=self.colors['primary'],
            width=100
        )
        decor_line.place(x=10, y=0, anchor=tk.NW)

        # 图片容器，居中显示
        img_container = tk.Frame(self.image_frame, bg=self.colors['light_bg'])
        img_container.pack(fill=tk.BOTH, expand=True)

        self.image_label = tk.Label(
            img_container,
            text="请选择图片",
            font=("微软雅黑", 12),
            bg=self.colors['light_bg'],
            fg=self.colors['text_light'],
            pady=50
        )
        self.image_label.pack(expand=True)

        # 结果显示区域 - 添加阴影和圆角效果
        result_frame = tk.LabelFrame(
            content_frame,
            text="识别结果 (LaTeX格式)",
            font=("微软雅黑", 10),
            padx=15,
            pady=15,
            bg=self.colors['light_bg'],
            relief=tk.FLAT,
            bd=1,
            highlightbackground=self.colors['border'],
            highlightthickness=1
        )
        result_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 添加简单的装饰线条
        decor_line = tk.Frame(
            result_frame,
            height=3,
            bg=self.colors['secondary'],
            width=100
        )
        decor_line.place(x=10, y=0, anchor=tk.NW)

        # 结果文本框样式优化
        self.result_text = scrolledtext.ScrolledText(
            result_frame,
            height=1,
            wrap=tk.WORD,
            font=("Consolas", 11),
            bg=self.colors['light_bg'],
            bd=0,
            highlightthickness=1,
            highlightbackground=self.colors['border']
        )
        self.result_text.pack(fill=tk.BOTH, expand=True)

        # 状态栏样式优化
        self.status_label = tk.Label(
            self.root,
            text="就绪",
            font=("微软雅黑", 9),
            bg=self.colors['light_bg'],
            fg=self.colors['text_light'],
            bd=1,
            relief=tk.FLAT,
            highlightbackground=self.colors['border'],
            highlightthickness=1,
            anchor=tk.W,
            padx=10,
            pady=5
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def load_model(self):
        """加载识别模型"""
        if self.model is None:
            self.status_label.config(text="正在加载模型...")
            self.root.update()

            # 显示加载动画
            self.show_loading("正在加载模型，请稍候...")

            try:
                # 确保环境变量正确设置
                cache_base = os.path.dirname(MODEL_CACHE_DIR)
                os.environ['TORCH_HOME'] = cache_base
                os.environ['HF_HOME'] = os.path.join(cache_base, "huggingface")
                os.environ['XDG_CACHE_HOME'] = cache_base
                os.environ['HUGGINGFACE_HUB_CACHE'] = os.path.join(cache_base, "huggingface", "hub")

                # 创建目录
                os.makedirs(os.path.join(cache_base, "models"), exist_ok=True)
                os.makedirs(os.path.join(cache_base, "huggingface"), exist_ok=True)

                print(f"模型缓存目录: {cache_base}")

                # 直接加载模型
                self.model = LatexOCR()

                # 隐藏加载动画
                self.hide_loading()

                self.status_label.config(text="模型加载完成", fg=self.colors['success'])
                return True

            except ImportError as e:
                # 隐藏加载动画
                self.hide_loading()

                error_msg = f"模型导入失败: {e}\n\n"
                if is_frozen():
                    error_msg += "请确保.exe文件包含了所有必要的依赖"
                else:
                    error_msg += "请运行: pip install pix2tex[gui]"

                messagebox.showerror("错误", error_msg)
                self.status_label.config(text="模型加载失败", fg=self.colors['error'])
                return False

            except FileNotFoundError as e:
                # 隐藏加载动画
                self.hide_loading()

                error_msg = f"文件路径错误: {e}\n\n"
                error_msg += f"模型缓存目录: {MODEL_CACHE_DIR}\n\n"
                error_msg += "解决方案:\n"
                error_msg += "1. 确保 _internal 文件夹与程序放在一起\n"
                error_msg += "2. 确保程序所在目录可写入缓存\n"
                error_msg += "3. 确保磁盘空间充足\n"
                if is_frozen():
                    error_msg += "4. 或者使用Python版本运行: python gui_version.py"

                messagebox.showerror("错误", error_msg)
                self.status_label.config(text="模型加载失败", fg=self.colors['error'])
                return False

            except Exception as e:
                # 隐藏加载动画
                self.hide_loading()

                import traceback
                error_details = traceback.format_exc()
                print(f"详细错误信息:\n{error_details}")

                error_msg = f"模型加载失败: {e}\n\n"
                error_msg += f"模型缓存目录: {MODEL_CACHE_DIR}\n\n"
                error_msg += "请检查:\n"
                error_msg += "1. _internal 文件夹是否完整\n"
                error_msg += "2. 程序所在目录是否可写\n"
                error_msg += "3. 磁盘空间是否充足\n"
                if is_frozen():
                    error_msg += "\n推荐: 使用Python版本运行更稳定\n"
                    error_msg += "运行: python gui_version.py"
                messagebox.showerror("错误", error_msg)
                self.status_label.config(text="模型加载失败", fg=self.colors['error'])
                return False
        return True

    def show_loading(self, message="正在处理..."):
        """显示加载动画"""
        # 如果已经有加载动画，先移除
        if hasattr(self, 'loading_window') and self.loading_window:
            self.hide_loading()

        self.loading = True

        # 创建加载窗口
        self.loading_window = tk.Toplevel(self.root)
        self.loading_window.title("")
        self.loading_window.geometry("300x150")
        self.loading_window.transient(self.root)
        self.loading_window.grab_set()  # 模态窗口
        self.loading_window.resizable(False, False)

        # 计算居中位置
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 150
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 75
        self.loading_window.geometry(f"+{x}+{y}")

        # 设置窗口样式
        self.loading_window.configure(bg=self.colors['light_bg'])

        # 添加加载图标（使用Unicode旋转符号）
        self.loading_label = tk.Label(
            self.loading_window,
            text="⟳",  # 旋转符号
            font=("Arial", 32),
            bg=self.colors['light_bg'],
            fg=self.colors['primary']
        )
        self.loading_label.pack(pady=20)

        # 添加加载文本
        text_label = tk.Label(
            self.loading_window,
            text=message,
            font=("微软雅黑", 11),
            bg=self.colors['light_bg'],
            fg=self.colors['foreground']
        )
        text_label.pack()

        # 开始旋转动画
        self.rotate_loading_icon()

    def rotate_loading_icon(self):
        """旋转加载图标"""
        if not self.loading or not hasattr(self, 'loading_label'):
            return

        # 使用不同的Unicode符号模拟旋转效果
        icons = ["⟳", "⟲", "⏳", "◷"]
        current_text = self.loading_label.cget("text")
        current_index = icons.index(current_text) if current_text in icons else 0
        next_index = (current_index + 1) % len(icons)

        self.loading_label.config(text=icons[next_index])

        # 继续旋转
        if self.loading:
            self.root.after(200, self.rotate_loading_icon)

    def hide_loading(self):
        """隐藏加载动画"""
        self.loading = False

        if hasattr(self, 'loading_window') and self.loading_window:
            try:
                self.loading_window.destroy()
            except:
                pass

            delattr(self, 'loading_window')

    def select_image(self):
        """选择图片文件"""
        file_path = filedialog.askopenfilename(
            title="选择图片",
            filetypes=[
                ("图片文件", "*.png *.jpg *.jpeg"),
                ("PNG文件", "*.png"),
                ("JPEG文件", "*.jpg *.jpeg"),
                ("所有文件", "*.*")
            ]
        )

        if file_path:
            self.current_image_path = file_path
            self.current_latex = None
            self.crop_box = None
            self.result_text.delete(1.0, tk.END)
            self.copy_latex_btn.config(state=tk.DISABLED)
            self.copy_img_btn.config(state=tk.DISABLED)
            self.export_btn.config(state=tk.DISABLED)
            self.crop_btn.config(state=tk.NORMAL)
            self.display_image(file_path)
            self.recognize_btn.config(state=tk.NORMAL)
            self.status_label.config(text=f"已选择: {os.path.basename(file_path)}")

    def select_formula_region(self):
        """Let the user exclude axes, captions and watermarks before OCR."""
        if not self.current_image_path:
            return
        with Image.open(self.current_image_path) as source:
            image = ImageOps.exif_transpose(source).copy()
        scale = min(900 / image.width, 600 / image.height, 1)
        preview = image.resize(
            (max(1, round(image.width * scale)), max(1, round(image.height * scale))),
            Image.LANCZOS,
        )

        dialog = tk.Toplevel(self.root)
        dialog.title("框选公式区域")
        dialog.transient(self.root)
        tk.Label(dialog, text="拖动鼠标框住公式，避开坐标轴、水印和其他文字").pack(pady=8)
        canvas = tk.Canvas(dialog, width=preview.width, height=preview.height,
                           cursor="crosshair", highlightthickness=0)
        canvas.pack(padx=10, pady=5)
        photo = ImageTk.PhotoImage(preview, master=dialog)
        canvas.create_image(0, 0, anchor=tk.NW, image=photo)
        canvas.image = photo
        selection = {'start': None, 'box': None, 'rectangle': None}

        def press(event):
            selection['start'] = (event.x, event.y)
            if selection['rectangle']:
                canvas.delete(selection['rectangle'])
            selection['rectangle'] = canvas.create_rectangle(
                event.x, event.y, event.x, event.y, outline='#E53935', width=2)

        def drag(event):
            if selection['start']:
                x, y = selection['start']
                canvas.coords(selection['rectangle'], x, y, event.x, event.y)

        def release(event):
            if not selection['start']:
                return
            x, y = selection['start']
            left, right = sorted((x, event.x))
            top, bottom = sorted((y, event.y))
            selection['box'] = (
                max(0, int(left / scale)), max(0, int(top / scale)),
                min(image.width, int(right / scale)),
                min(image.height, int(bottom / scale)),
            )
            selection['start'] = None

        def apply_crop(box):
            if box is not None and (box[2] - box[0] < 10 or box[3] - box[1] < 10):
                messagebox.showwarning("提示", "请框选完整的公式", parent=dialog)
                return
            self.crop_box = box
            self.current_latex = None
            self.result_text.delete(1.0, tk.END)
            self.copy_latex_btn.config(state=tk.DISABLED)
            self.copy_img_btn.config(state=tk.DISABLED)
            self.export_btn.config(state=tk.DISABLED)
            self.display_image(self.current_image_path)
            self.status_label.config(text="已框选公式，请点击“识别公式”" if box else "已恢复整张图片")
            dialog.destroy()

        canvas.bind('<ButtonPress-1>', press)
        canvas.bind('<B1-Motion>', drag)
        canvas.bind('<ButtonRelease-1>', release)
        controls = tk.Frame(dialog)
        controls.pack(pady=8)
        def apply_selected():
            if selection['box'] is None:
                messagebox.showwarning("提示", "请先拖动鼠标框选公式", parent=dialog)
            else:
                apply_crop(selection['box'])

        ttk.Button(controls, text="应用区域",
                   command=apply_selected).pack(side=tk.LEFT, padx=8)
        ttk.Button(controls, text="使用整张图",
                   command=lambda: apply_crop(None)).pack(side=tk.LEFT, padx=8)
        ttk.Button(controls, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=8)

    def display_image(self, image_path):
        """显示图片预览"""
        try:
            with Image.open(image_path) as source:
                img = ImageOps.exif_transpose(source).copy()
            if self.crop_box:
                img = img.crop(self.crop_box)

            # 获取容器大小，计算合适的图片尺寸
            container_width = self.image_frame.winfo_width() - 40  # 减去padding
            container_height = self.image_frame.winfo_height() - 40  # 减去padding

            # 如果容器大小还未确定（首次加载），使用默认值
            if container_width < 100 or container_height < 100:
                container_width, container_height = 400, 300

            # 调整图片大小以适应显示区域，保持原始比例
            img.thumbnail((container_width, container_height), Image.LANCZOS)

            photo = ImageTk.PhotoImage(img)
            self.image_label.config(
                image=photo,
                text="",
                pady=0,
                highlightthickness=1,
                highlightbackground=self.colors['border']
            )
            self.image_label.image = photo  # 保持引用

            # 绑定窗口大小变化事件，重新调整图片大小
            self.root.bind('<Configure>', self._on_resize)

        except Exception as e:
            messagebox.showerror("错误", f"无法显示图片: {e}")

    def _on_resize(self, event):
        """窗口大小变化时重新调整图片大小"""
        if hasattr(self, 'current_image_path') and self.current_image_path:
            # 避免频繁更新，添加延迟
            self.root.after(200, lambda: self.display_image(self.current_image_path))

    def recognize_formula(self):
        """识别公式"""
        if not self.current_image_path:
            messagebox.showwarning("警告", "请先选择图片")
            return

        if not self.load_model():
            return

        # 禁用识别按钮，防止重复点击
        self.recognize_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        self.crop_btn.config(state=tk.DISABLED)
        self.status_label.config(text="正在识别...")
        self.root.update()

        # 显示加载动画
        self.show_loading("正在识别公式，请稍候...")

        # 在后台线程中执行识别，避免阻塞GUI
        image_path = self.current_image_path
        crop_box = self.crop_box

        def recognize_thread():
            try:
                # Respect camera rotation and place transparent screenshots on white.
                with Image.open(image_path) as source:
                    img = ImageOps.exif_transpose(source).convert('RGBA')
                if crop_box:
                    img = img.crop(crop_box)
                background = Image.new('RGBA', img.size, 'white')
                img = Image.alpha_composite(background, img).convert('RGB')
                latex_formula = self.model(img)

                # 快速检查识别结果
                if latex_formula is None:
                    raise Exception("识别结果为空")

                # 确保是字符串类型
                latex_formula = str(latex_formula).strip()

                if not latex_formula:
                    raise Exception("识别结果为空字符串")

                # 隐藏加载动画
                self.root.after(0, self.hide_loading)

                # 在主线程中更新UI
                self.root.after(0, lambda: self._recognition_completed(latex_formula))

            except Exception as e:
                import traceback
                traceback.print_exc()

                # 隐藏加载动画
                self.root.after(0, self.hide_loading)

                # 在主线程中显示错误
                self.root.after(0, lambda message=str(e): self._recognition_failed(message))

        # 启动后台线程
        thread = threading.Thread(target=recognize_thread, daemon=True)
        thread.start()

    def _recognition_completed(self, latex_formula):
        """识别完成回调"""
        try:
            self.current_latex = latex_formula

            # 允许用户在复制前纠正 OCR 的误识别。
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(1.0, latex_formula)

            self.copy_latex_btn.config(state=tk.NORMAL)
            self.copy_img_btn.config(state=tk.NORMAL)
            self.export_btn.config(state=tk.NORMAL)
            self.recognize_btn.config(state=tk.NORMAL)
            self.select_btn.config(state=tk.NORMAL)
            self.crop_btn.config(state=tk.NORMAL)

            self.status_label.config(text="识别完成", fg=self.colors['success'])
            messagebox.showinfo("成功", "公式识别完成！")

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.recognize_btn.config(state=tk.NORMAL)
            self.select_btn.config(state=tk.NORMAL)
            self.crop_btn.config(state=tk.NORMAL)
            messagebox.showerror("错误", f"更新结果失败: {e}")
            self.status_label.config(text="识别失败")

    def _recognition_failed(self, error_msg):
        """识别失败回调"""
        self.recognize_btn.config(state=tk.NORMAL)
        self.select_btn.config(state=tk.NORMAL)
        self.crop_btn.config(state=tk.NORMAL)
        self.status_label.config(text="识别失败", fg=self.colors['error'])
        messagebox.showerror("错误", f"识别失败: {error_msg}\n\n请检查:\n1. 图片是否清晰\n2. 图片是否包含数学公式\n3. 模型是否正确加载")

    def copy_latex(self):
        """复制LaTeX到剪贴板"""
        if self.current_latex is not None:
            try:
                # Use the editable text box, including any user corrections.
                cleaned_latex = self.clean_latex(self.result_text.get(1.0, tk.END))
                if not cleaned_latex:
                    messagebox.showwarning("提示", "识别结果为空，请先输入或识别公式")
                    return

                pyperclip.copy(cleaned_latex)

                self.status_label.config(text="LaTeX已复制到剪贴板", fg=self.colors['success'])

                # 创建更美观的提示框
                msg_window = tk.Toplevel(self.root)
                msg_window.title("复制成功")
                msg_window.geometry("400x200")
                msg_window.transient(self.root)
                msg_window.resizable(False, False)

                # 计算居中位置
                x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
                y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 100
                msg_window.geometry(f"+{x}+{y}")

                # 设置消息框样式
                msg_window.configure(bg=self.colors['light_bg'])

                # 图标
                icon_label = tk.Label(msg_window, text="✓", font=("Arial", 24), bg=self.colors['light_bg'], fg=self.colors['success'])
                icon_label.pack(pady=10)

                # 消息
                msg_label = tk.Label(
                    msg_window,
                    text="LaTeX公式已复制到剪贴板！",
                    font=("微软雅黑", 11),
                    bg=self.colors['light_bg'],
                    fg=self.colors['foreground']
                )
                msg_label.pack(pady=5)

                # 使用提示
                hint_label = tk.Label(
                    msg_window,
                    text="Word：Alt+= → 选择LaTeX → 粘贴 → Ctrl+= 转换",
                    font=("微软雅黑", 10),
                    bg=self.colors['light_bg'],
                    fg=self.colors['text_light']
                )
                hint_label.pack(pady=5)

                # 确定按钮
                ok_btn = ttk.Button(
                    msg_window,
                    text="确定",
                    command=msg_window.destroy,
                    style="Modern.TButton",
                    width=10
                )
                ok_btn.pack(pady=10)

                # 绑定回车键关闭
                msg_window.bind('<Return>', lambda e: msg_window.destroy())
                ok_btn.focus_set()

            except Exception as e:
                messagebox.showerror("错误", f"复制失败: {e}")
                self.status_label.config(text="复制失败", fg=self.colors['error'])

    def copy_image(self):
        """复制图片到剪贴板"""
        if not self.current_image_path:
            messagebox.showwarning("警告", "请先选择图片")
            return

        try:
            success = copy_image_to_clipboard(self.current_image_path, self.crop_box)
            if success:
                self.status_label.config(text="图片已复制到剪贴板")
                messagebox.showinfo(
                    "成功",
                    "图片已复制到剪贴板！\n\n"
                    "现在可以直接在Word中粘贴（Ctrl+V）"
                )
            else:
                raise Exception("复制操作失败")
        except Exception as e:
            messagebox.showerror("错误", f"复制图片失败: {e}")
            self.status_label.config(text="复制失败")

    def export_word_formula(self):
        """Save the corrected OCR result as a native, editable Word equation."""
        formula = self.clean_latex(self.result_text.get(1.0, tk.END))
        if not formula:
            messagebox.showwarning("提示", "识别结果为空")
            return
        output_path = filedialog.asksaveasfilename(
            title="保存可编辑Word公式",
            defaultextension=".docx",
            filetypes=[("Word文档", "*.docx")],
            initialfile="formula_output.docx",
        )
        if not output_path:
            return
        try:
            doc = Document()
            if not insert_latex_to_word(doc, formula):
                raise ValueError("公式转换失败，请检查LaTeX内容及Microsoft Office安装")
            doc.save(output_path)
            self.status_label.config(text=f"已保存Word公式: {output_path}", fg=self.colors['success'])
            messagebox.showinfo("保存成功", "已生成可编辑Word公式，可从文档中复制到目标文档。")
        except Exception as error:
            messagebox.showerror("保存失败", str(error))




def main():
    """主函数"""
    root = tk.Tk()
    app = FormulaRecognizerGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
