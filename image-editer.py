import os
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import Image, ImageTk
import imghdr
from docx import Document
import re


class ZoomableImage(ttk.Frame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.image = None
        self.photo_image = None
        self.canvas = tk.Canvas(self, bg='white', highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # 初始化缩放参数
        self.scale = 1.0
        self.max_scale = 4.0
        self.min_scale = 0.1

        # 初始化拖动参数
        self.x = 0
        self.y = 0
        self.last_x = 0
        self.last_y = 0

        # 绑定事件
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_move_press)
        self.canvas.bind("<Configure>", self.on_canvas_resize)

    def set_image(self, image_path):
        try:
            if image_path and os.path.exists(image_path):
                self.image = Image.open(image_path)
                self.reset_view()
                self.update_image()
                return True
            else:
                self.show_message("无图片可显示")
                return False
        except Exception as e:
            self.show_message("图片加载失败")
            print(f"图片加载错误: {e}")  # 在控制台记录错误，但不弹出对话框
            return False

    def show_message(self, message):
        """在画布上显示消息"""
        self.image = None
        self.photo_image = None
        self.canvas.delete("all")

        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        if canvas_width > 1 and canvas_height > 1:
            self.canvas.create_text(
                canvas_width // 2,
                canvas_height // 2,
                text=message,
                font=('Arial', 16, 'bold'),
                fill="gray",
                width=canvas_width - 40
            )

    def reset_view(self):
        self.scale = 1.0
        self.x = 0
        self.y = 0
        self.update_image()

    def update_image(self):
        if not self.image:
            return

        self._apply_transform()

    def _apply_transform(self):
        if not self.image:
            return

        self.canvas.delete("all")

        # 计算缩放后的尺寸
        img_width, img_height = self.image.size
        scaled_width = int(img_width * self.scale)
        scaled_height = int(img_height * self.scale)

        # 调整图片位置（居中或根据拖动位置）
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # 限制拖动范围
        max_x = max(0, (scaled_width - canvas_width) // 2)
        max_y = max(0, (scaled_height - canvas_height) // 2)
        self.x = max(-max_x, min(max_x, self.x))
        self.y = max(-max_y, min(max_y, self.y))

        # 创建缩放后的图片
        resized_image = self.image.resize((scaled_width, scaled_height), Image.Resampling.LANCZOS)
        self.photo_image = ImageTk.PhotoImage(resized_image)

        # 在画布上显示图片
        x_pos = canvas_width // 2 + self.x
        y_pos = canvas_height // 2 + self.y
        self.canvas.create_image(x_pos, y_pos, image=self.photo_image, anchor=tk.CENTER)

        # 显示缩放比例
        self.canvas.create_text(
            10, 10,
            text=f"缩放: {self.scale * 100:.0f}%",
            anchor=tk.NW,
            fill="black",
            font=('Arial', 10, 'bold')
        )

    def zoom(self, factor, x=None, y=None):
        if not self.image:
            return

        old_scale = self.scale
        self.scale *= factor
        self.scale = max(self.min_scale, min(self.max_scale, self.scale))

        if x is not None and y is not None and old_scale != 0:
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()

            rel_x = x - canvas_width // 2
            rel_y = y - canvas_height // 2

            self.x = rel_x - (rel_x - self.x) * (self.scale / old_scale)
            self.y = rel_y - (rel_y - self.y) * (self.scale / old_scale)

        self._apply_transform()

    def on_mouse_wheel(self, event):
        if self.image:  # 只在有图片时响应缩放
            factor = 1.1 if event.delta > 0 else 0.9
            self.zoom(factor, event.x, event.y)

    def on_button_press(self, event):
        if self.image:  # 只在有图片时响应拖动
            self.last_x = event.x
            self.last_y = event.y

    def on_move_press(self, event):
        if self.image:  # 只在有图片时响应拖动
            dx = event.x - self.last_x
            dy = event.y - self.last_y
            self.last_x = event.x
            self.last_y = event.y

            self.x += dx
            self.y += dy
            self._apply_transform()

    def on_canvas_resize(self, event):
        self._apply_transform()


class WordImageExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word图片提取与重命名工具")
        self.root.geometry("1000x800")

        # 初始化变量
        self.word_file_path = tk.StringVar()
        self.image_folder_path = tk.StringVar()
        self.image_files = []
        self.current_index = 0
        self.current_image_path = ""

        # 创建界面
        self.create_widgets()

        # 绑定键盘事件
        self.root.bind('<Left>', lambda e: self.previous_image())
        self.root.bind('<Right>', lambda e: self.next_image())
        self.root.bind('<Up>', lambda e: self.previous_image())
        self.root.bind('<Down>', lambda e: self.next_image())
        self.root.focus_set()  # 确保窗口可以接收键盘事件

    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建Notebook（选项卡）
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 创建两个选项卡
        self.create_extract_tab()
        self.create_rename_tab()

    def create_extract_tab(self):
        """创建图片提取选项卡"""
        extract_tab = ttk.Frame(self.notebook)
        self.notebook.add(extract_tab, text="从Word提取图片")

        # Word文件选择区域
        file_frame = ttk.LabelFrame(extract_tab, text="Word文件选择", padding="10")
        file_frame.pack(fill=tk.X, pady=5)

        ttk.Label(file_frame, text="Word文件路径:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(file_frame, textvariable=self.word_file_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="浏览...", command=self.browse_word_file).grid(row=0, column=2)

        # 输出目录选择区域
        output_frame = ttk.LabelFrame(extract_tab, text="输出目录", padding="10")
        output_frame.pack(fill=tk.X, pady=5)

        ttk.Label(output_frame, text="输出文件夹路径:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(output_frame, textvariable=self.image_folder_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(output_frame, text="浏览...", command=self.browse_output_folder).grid(row=0, column=2)

        # 操作按钮区域
        button_frame = ttk.Frame(extract_tab)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="提取图片", command=self.extract_images).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="打开图片文件夹", command=self.open_image_folder).pack(side=tk.LEFT, padx=5)

    def create_rename_tab(self):
        """创建图片重命名选项卡"""
        rename_tab = ttk.Frame(self.notebook)
        self.notebook.add(rename_tab, text="图片重命名")

        # 图片文件夹选择区域
        folder_frame = ttk.LabelFrame(rename_tab, text="图片文件夹", padding="10")
        folder_frame.pack(fill=tk.X, pady=5)

        ttk.Label(folder_frame, text="图片文件夹路径:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(folder_frame, textvariable=self.image_folder_path, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(folder_frame, text="浏览...", command=self.browse_image_folder).grid(row=0, column=2)

        ttk.Button(folder_frame, text="加载图片", command=self.load_image_files).grid(row=1, column=2, pady=5,
                                                                                      sticky=tk.E)

        # 图片显示与重命名区域
        image_frame = ttk.LabelFrame(rename_tab, text="图片预览与重命名", padding="10")
        image_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # 使用自定义的可缩放图片组件
        self.zoomable_image = ZoomableImage(image_frame)
        self.zoomable_image.pack(fill=tk.BOTH, expand=True)

        # 文件名编辑区域
        control_frame = ttk.Frame(image_frame)
        control_frame.pack(fill=tk.X, pady=5)

        # 文件名编辑
        name_frame = ttk.Frame(control_frame)
        name_frame.pack(side=tk.LEFT, padx=5)

        ttk.Label(name_frame, text="文件名:").pack(side=tk.LEFT)
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(name_frame, textvariable=self.name_var, width=30)
        self.name_entry.pack(side=tk.LEFT, padx=5)
        self.name_entry.bind('<Return>', self.save_rename)

        # 文件扩展名显示
        self.ext_var = tk.StringVar()
        ttk.Label(name_frame, textvariable=self.ext_var).pack(side=tk.LEFT, padx=5)

        # 缩放控制按钮
        zoom_frame = ttk.Frame(control_frame)
        zoom_frame.pack(side=tk.RIGHT, padx=5)

        ttk.Button(zoom_frame, text="放大", command=lambda: self.zoomable_image.zoom(1.2)).pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="缩小", command=lambda: self.zoomable_image.zoom(0.8)).pack(side=tk.LEFT, padx=2)
        ttk.Button(zoom_frame, text="重置", command=self.zoomable_image.reset_view).pack(side=tk.LEFT, padx=2)

        # 导航按钮
        nav_frame = ttk.Frame(control_frame)
        nav_frame.pack(side=tk.RIGHT, padx=10)

        ttk.Button(nav_frame, text="上一张", command=self.previous_image).pack(side=tk.LEFT, padx=2)
        ttk.Button(nav_frame, text="下一张", command=self.next_image).pack(side=tk.LEFT, padx=2)
        ttk.Button(nav_frame, text="保存重命名", command=self.save_rename).pack(side=tk.LEFT, padx=2)

        # 初始状态显示提示
        self.zoomable_image.show_message("请选择图片文件夹并加载图片")

    def browse_word_file(self):
        file_path = filedialog.askopenfilename(
            title="选择 Word 文件",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.word_file_path.set(file_path)
            # 自动设置输出目录为 Word 文件所在目录下的 images 文件夹
            word_dir = os.path.dirname(file_path)
            default_output = os.path.join(word_dir, "extracted_images")
            self.image_folder_path.set(default_output)

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory(title="选择输出文件夹")
        if folder_path:
            self.image_folder_path.set(folder_path)

    def browse_image_folder(self):
        folder_path = filedialog.askdirectory(title="选择图片文件夹")
        if folder_path:
            self.image_folder_path.set(folder_path)

    def extract_images(self):
        word_file = self.word_file_path.get()
        output_folder = self.image_folder_path.get()

        if not word_file or not output_folder:
            messagebox.showerror("错误", "请先选择 Word 文件和输出文件夹")
            return

        if not word_file.endswith('.docx'):
            messagebox.showerror("错误", "请选择有效的 .docx 文件")
            return

        try:
            # 确保输出文件夹存在
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            # 获取图片在文档中的实际顺序
            image_order = self.get_image_order_from_docx(word_file)

            # 使用zip解压获取图片文件
            with zipfile.ZipFile(word_file, 'r') as docx_zip:
                # 获取所有媒体文件
                media_files = [f for f in docx_zip.namelist() if f.startswith('word/media/')]

                # 清空目标文件夹（可选）
                for existing_file in os.listdir(output_folder):
                    file_path = os.path.join(output_folder, existing_file)
                    try:
                        if os.path.isfile(file_path):
                            os.unlink(file_path)
                    except Exception as e:
                        print(f"删除文件 {file_path} 失败: {e}")

                # 按检测到的顺序提取图片
                valid_images = []
                for i, rel_path in enumerate(image_order, 1):
                    if rel_path in media_files:
                        # 从zip文件中读取图片数据
                        with docx_zip.open(rel_path) as source:
                            image_data = source.read()

                        # 检测图片实际类型
                        image_type = imghdr.what(None, h=image_data)
                        if not image_type:
                            continue  # 不是有效图片，跳过

                        # 确定文件扩展名
                        ext_map = {
                            'jpeg': '.jpg',
                            'jpg': '.jpg',
                            'png': '.png',
                            'bmp': '.bmp',
                            'gif': '.gif',
                            'tiff': '.tiff',
                            'webp': '.webp'
                        }
                        ext = ext_map.get(image_type, '.png')

                        # 新文件名
                        new_filename = f"{i:03d}{ext}"
                        output_path = os.path.join(output_folder, new_filename)

                        # 保存图片
                        with open(output_path, 'wb') as target:
                            target.write(image_data)

                        valid_images.append(new_filename)

            if valid_images:
                messagebox.showinfo("完成", f"成功提取 {len(valid_images)} 张图片到 {output_folder}")
                # 自动切换到重命名标签页
                self.notebook.select(1)
                self.load_image_files(output_folder)
            else:
                messagebox.showwarning("警告", "未找到有效的图片文件")

        except Exception as e:
            messagebox.showerror("错误", f"提取图片失败: {str(e)}")

    def get_image_order_from_docx(self, docx_path):
        """通过解析document.xml获取图片在文档中的实际顺序"""
        image_order = []

        # 方法1：使用python-docx获取图片引用
        try:
            doc = Document(docx_path)
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    image_path = f"word/{rel.target_ref}"
                    image_order.append(image_path)
        except:
            pass

        # 方法2：直接解析document.xml（更准确）
        try:
            with zipfile.ZipFile(docx_path) as z:
                with z.open('word/document.xml') as f:
                    xml_content = f.read().decode('utf-8')

            # 查找所有图片引用
            image_refs = re.findall(r'<a:blip r:embed="([^"]+)"', xml_content)
            for ref in image_refs:
                image_path = f"word/{ref}"
                image_order.append(image_path)
        except:
            pass

        # 如果以上方法都失败，回退到按文件名排序
        if not image_order:
            with zipfile.ZipFile(docx_path) as z:
                media_files = [f for f in z.namelist() if f.startswith('word/media/')]
                image_order = sorted(media_files)

        return image_order

    def open_image_folder(self):
        output_folder = self.image_folder_path.get()
        if not output_folder:
            messagebox.showerror("错误", "请先选择输出文件夹")
            return

        if not os.path.exists(output_folder):
            messagebox.showerror("错误", "输出文件夹不存在，请先提取图片")
            return

        try:
            # 打开文件夹
            if os.name == 'nt':  # Windows
                os.startfile(output_folder)
            elif os.name == 'mac':  # macOS
                os.system(f'open "{output_folder}"')
            else:  # Linux
                os.system(f'xdg-open "{output_folder}"')
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件夹: {str(e)}")

    def load_image_files(self, folder_path=None):
        if not folder_path:
            folder_path = self.image_folder_path.get()

        if not folder_path:
            messagebox.showerror("错误", "请先选择图片文件夹")
            return

        if not os.path.exists(folder_path):
            messagebox.showerror("错误", "图片文件夹不存在")
            return

        # 获取文件夹中的所有图片文件（包括无扩展名的）
        self.image_files = []
        for f in os.listdir(folder_path):
            file_path = os.path.join(folder_path, f)
            if os.path.isfile(file_path):
                # 检测文件是否为图片
                try:
                    image_type = imghdr.what(file_path)
                    if image_type:
                        self.image_files.append(f)
                except:
                    continue

        # 按数字序号排序
        self.image_files.sort(key=lambda x: int(''.join(filter(str.isdigit, x)) or '0'))

        if self.image_files:
            self.current_index = 0
            self.show_image()
        else:
            self.show_no_images_message()

    def show_image(self):
        if not self.image_files:
            self.zoomable_image.show_message("无图片可显示")
            self.clear_file_info()
            return

        if 0 <= self.current_index < len(self.image_files):
            folder_path = self.image_folder_path.get()
            image_file = self.image_files[self.current_index]
            self.current_image_path = os.path.join(folder_path, image_file)

            # 尝试加载图片，如果失败则显示友好消息
            success = self.zoomable_image.set_image(self.current_image_path)
            if success:
                try:
                    # 更新文件名和扩展名显示
                    _, ext = os.path.splitext(image_file)
                    self.ext_var.set(ext)

                    # 只显示文件名部分（不含扩展名）
                    base_name, _ = os.path.splitext(image_file)
                    self.name_var.set(base_name)
                except Exception as e:
                    print(f"文件信息更新错误: {e}")
                    self.clear_file_info()
            else:
                self.clear_file_info()
        else:
            self.show_completion_message()

    def clear_file_info(self):
        """清空文件信息显示"""
        self.name_var.set("")
        self.ext_var.set("")

    def show_no_images_message(self):
        """显示无图片消息"""
        self.zoomable_image.show_message("文件夹中没有找到图片文件")
        self.clear_file_info()
        self.current_index = 0
        self.current_image_path = ""

    def show_completion_message(self):
        """显示查看完成消息"""
        if self.image_files:
            self.zoomable_image.show_message("图片已查看完毕")
        else:
            self.zoomable_image.show_message("无图片可显示")
        self.clear_file_info()

    def previous_image(self):
        if self.image_files:
            if self.current_index > 0:
                self.current_index -= 1
                self.show_image()
            else:
                self.show_completion_message()
        else:
            self.show_no_images_message()

    def next_image(self):
        if self.image_files:
            if self.current_index < len(self.image_files) - 1:
                self.current_index += 1
                self.show_image()
            else:
                self.show_completion_message()
        else:
            self.show_no_images_message()

    def save_rename(self, event=None):
        if not self.current_image_path or not self.image_files:
            messagebox.showwarning("警告", "没有可重命名的图片")
            return

        new_base_name = self.name_var.get().strip()
        if not new_base_name:
            messagebox.showerror("错误", "文件名不能为空")
            return

        # 获取原文件的扩展名
        _, ext = os.path.splitext(self.current_image_path)
        if not ext:
            # 如果没有扩展名，尝试检测
            image_type = imghdr.what(self.current_image_path)
            ext_map = {
                'jpeg': '.jpg',
                'jpg': '.jpg',
                'png': '.png',
                'bmp': '.bmp',
                'gif': '.gif',
                'tiff': '.tiff',
                'webp': '.webp'
            }
            ext = ext_map.get(image_type, '.png')

        new_name = new_base_name + ext
        folder_path = self.image_folder_path.get()
        new_path = os.path.join(folder_path, new_name)

        if self.current_image_path == new_path:
            return  # 没有变化

        try:
            # 检查新文件名是否已存在
            if os.path.exists(new_path):
                messagebox.showerror("错误", "文件名已存在")
                return

            # 重命名文件
            os.rename(self.current_image_path, new_path)

            # 更新文件列表和当前路径
            self.image_files[self.current_index] = new_name
            self.current_image_path = new_path

            messagebox.showinfo("成功", "文件名已更新")
        except Exception as e:
            messagebox.showerror("错误", f"重命名失败: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = WordImageExtractorApp(root)
    root.mainloop()