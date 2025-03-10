import os
import logging
from tkinter import Tk, Button, Label, filedialog, messagebox, StringVar, Radiobutton, Listbox, Scrollbar, Frame, LEFT, RIGHT, Y, BOTH, END, Menu, Toplevel
from tkinter.ttk import Notebook
from PIL import Image
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
import docx2pdf  # 用于 Word 转 PDF
import win32com.client  # 用于 PPT 转 PDF
from pdftofileclass import PDFToFileConverter

# 配置日志记录
logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

# 定义一些颜色常量，方便后续使用
BG_COLOR = "#f0f0f0"
BUTTON_COLOR = "#e0e0e0"
TEXT_COLOR = "#333333"
TITLE_COLOR = "#222222"
FRAME_BG_COLOR = "#ffffff"


class ImageToPDFConverter:
    def __init__(self, tab):
        self.tab = tab

        # 创建变量存储选择的图片路径和输出文件夹路径
        self.image_paths_var = StringVar()
        self.output_folder_var = StringVar()
        self.page_orientation_var = StringVar()
        self.page_orientation_var.set("竖屏")  # 默认选择竖屏
        self.output_mode_var = StringVar()
        self.output_mode_var.set("默认")  # 默认选择默认输出

        # 创建标签
        self.status_label = Label(tab, text="请选择图片文件", wraplength=380, bg=BG_COLOR, fg=TEXT_COLOR,
                                  font=("Arial", 12))
        self.status_label.pack(pady=10)

        # 创建左侧框架用于放置按钮
        left_frame = Frame(tab, bg=FRAME_BG_COLOR, bd=2, relief="groove")
        left_frame.pack(side=LEFT, padx=10, pady=10, fill=Y)

        # 创建选择图片按钮
        select_button = Button(left_frame, text="选择图片", command=self.select_images, bg=BUTTON_COLOR,
                               fg=TEXT_COLOR, font=("Arial", 10))
        select_button.pack(pady=5, padx=10, fill=BOTH)

        # 创建选择输出文件夹按钮
        select_folder_button = Button(left_frame, text="选择输出文件夹", command=self.select_output_folder,
                                      bg=BUTTON_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        select_folder_button.pack(pady=5, padx=10, fill=BOTH)

        # 创建页面方向选择 Radiobutton 控件
        orientation_label = Label(left_frame, text="选择页面方向:", bg=FRAME_BG_COLOR, fg=TITLE_COLOR,
                                  font=("Arial", 11))
        orientation_label.pack(pady=10, padx=10)

        radio_vertical = Radiobutton(left_frame, text="竖屏", variable=self.page_orientation_var, value="竖屏",
                                     bg=FRAME_BG_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        radio_vertical.pack(pady=2, padx=10, anchor="w")

        radio_horizontal = Radiobutton(left_frame, text="横屏", variable=self.page_orientation_var, value="横屏",
                                       bg=FRAME_BG_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        radio_horizontal.pack(pady=2, padx=10, anchor="w")

        # 创建输出模式选择 Radiobutton 控件
        output_mode_label = Label(left_frame, text="选择输出模式:", bg=FRAME_BG_COLOR, fg=TITLE_COLOR,
                                  font=("Arial", 11))
        output_mode_label.pack(pady=10, padx=10)

        radio_default = Radiobutton(left_frame, text="默认", variable=self.output_mode_var, value="默认",
                                    bg=FRAME_BG_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        radio_default.pack(pady=2, padx=10, anchor="w")

        radio_fullscreen = Radiobutton(left_frame, text="全屏输出", variable=self.output_mode_var, value="全屏输出",
                                       bg=FRAME_BG_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        radio_fullscreen.pack(pady=2, padx=10, anchor="w")

        # 创建生成 PDF 按钮
        generate_button = Button(left_frame, text="生成 PDF", command=self.generate_pdf, bg=BUTTON_COLOR,
                                 fg=TEXT_COLOR, font=("Arial", 10))
        generate_button.pack(pady=15, padx=10, fill=BOTH)

        # 创建右侧框架用于放置 Listbox
        right_frame = Frame(tab, bg=FRAME_BG_COLOR, bd=2, relief="groove")
        right_frame.pack(side=RIGHT, padx=10, pady=10, fill=BOTH, expand=True)

        # 创建 Listbox 用于展示图片列表
        self.image_listbox = Listbox(right_frame, bg=BG_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        self.image_listbox.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)

        # 创建滚动条
        scrollbar = Scrollbar(right_frame, command=self.image_listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.image_listbox.config(yscrollcommand=scrollbar.set)

    def images_to_pdf(self, image_paths, pdf_path, page_orientation, output_mode):
        # 根据页面方向设置页面大小
        if page_orientation == "竖屏":
            page_size = A4
        else:
            page_size = landscape(A4)

        # 创建 PDF 文件
        try:
            c = canvas.Canvas(pdf_path, pagesize=page_size)

            for image_path in image_paths:
                try:
                    # 打开图片
                    img = Image.open(image_path)

                    # 获取图片的宽度和高度
                    img_width, img_height = img.size

                    # 获取 PDF 页面的宽度和高度
                    pdf_width, pdf_height = page_size

                    if output_mode == "全屏输出":
                        # 全屏输出，直接将图片铺满页面
                        new_width = pdf_width
                        new_height = pdf_height
                        x = 0
                        y = 0
                    else:
                        # 默认输出，计算缩放比例，以适应整个页面并保持原有比例
                        scale = min(pdf_width / img_width, pdf_height / img_height)
                        new_width = img_width * scale
                        new_height = img_height * scale

                        # 计算图片在 PDF 页面上的位置，使其居中显示
                        x = (pdf_width - new_width) / 2
                        y = (pdf_height - new_height) / 2

                    # 将图片绘制到 PDF 页面上
                    c.drawImage(image_path, x, y, width=new_width, height=new_height)

                    # 添加分页符
                    c.showPage()
                except Exception as e:
                    messagebox.showerror("错误", f"处理图片 {image_path} 时出错: {e}")
                    logging.error(f"处理图片 {image_path} 时出错: {e}")
            # 保存 PDF 文件
            c.save()
        except Exception as e:
            messagebox.showerror("错误", f"生成 PDF 文件时出错: {e}")
            logging.error(f"生成 PDF 文件时出错: {e}")

    def select_images(self):
        # 打开文件选择对话框，允许多选图片文件
        image_paths = filedialog.askopenfilenames(filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif")])
        if image_paths:
            for path in image_paths:
                self.image_listbox.insert(END, path)
            self.image_paths_var.set(",".join(image_paths))
            self.status_label.config(text=f"已选择 {len(image_paths)} 张图片")
        else:
            self.status_label.config(text="未选择任何图片")

    def select_output_folder(self):
        # 打开文件夹选择对话框
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_var.set(folder_path)
            self.status_label.config(text=f"输出文件夹已选择: {folder_path}")
        else:
            self.status_label.config(text="未选择输出文件夹")

    def generate_pdf(self):
        selected_paths = list(self.image_listbox.get(0, END))
        if not selected_paths:
            messagebox.showwarning("警告", "请先选择图片文件")
            return

        folder_path = self.output_folder_var.get()
        if not folder_path:
            messagebox.showwarning("警告", "请先选择输出文件夹")
            return

        # 使用 asksaveasfilename 选择文件名和路径
        pdf_filename = filedialog.asksaveasfilename(
            initialdir=folder_path,
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="output.pdf"
        )

        if pdf_filename:
            try:
                # 确保路径格式正确
                pdf_path = os.path.normpath(pdf_filename)
                print(f"生成的 PDF 路径: {pdf_path}")  # 调试信息

                # 检查路径是否有效
                if not os.path.isdir(os.path.dirname(pdf_path)):
                    raise FileNotFoundError(f"目录不存在: {os.path.dirname(pdf_path)}")

                # 获取用户选择的页面方向
                page_orientation = self.page_orientation_var.get()
                # 获取用户选择的输出模式
                output_mode = self.output_mode_var.get()

                self.images_to_pdf(selected_paths, pdf_path, page_orientation, output_mode)
                messagebox.showinfo("成功", f"PDF 文件已生成: {pdf_path}")
                self.status_label.config(text=f"PDF 文件已生成: {pdf_path}")
            except FileNotFoundError as e:
                messagebox.showerror("错误", f"文件或目录不存在: {e}")
                logging.error(f"文件或目录不存在: {e}")
            except PermissionError as e:
                messagebox.showerror("错误", f"权限不足: {e}")
                logging.error(f"权限不足: {e}")
            except Exception as e:
                messagebox.showerror("错误", f"生成 PDF 文件时出错: {e}")
                logging.error(f"生成 PDF 文件时出错: {e}")


class ExcelToPDFConverter:
    def __init__(self, tab):
        self.tab = tab

        # 创建变量存储选择的 Excel 文件路径和输出文件夹路径
        self.file_paths_var = StringVar()
        self.output_folder_var = StringVar()

        # 创建标签
        self.status_label = Label(tab, text="请选择 Excel 文件", wraplength=380, bg=BG_COLOR, fg=TEXT_COLOR,
                                  font=("Arial", 12))
        self.status_label.pack(pady=10)

        # 创建左侧框架用于放置按钮
        left_frame = Frame(tab, bg=FRAME_BG_COLOR, bd=2, relief="groove")
        left_frame.pack(side=LEFT, padx=10, pady=10, fill=Y)

        # 创建选择 Excel 文件按钮
        select_file_button = Button(left_frame, text="选择 Excel 文件", command=self.select_files, bg=BUTTON_COLOR,
                                    fg=TEXT_COLOR, font=("Arial", 10))
        select_file_button.pack(pady=5, padx=10, fill=BOTH)

        # 创建选择输出文件夹按钮
        select_folder_button = Button(left_frame, text="选择输出文件夹", command=self.select_output_folder,
                                      bg=BUTTON_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        select_folder_button.pack(pady=5, padx=10, fill=BOTH)

        # 创建生成 PDF 按钮
        generate_button = Button(left_frame, text="生成 PDF", command=self.generate_pdf, bg=BUTTON_COLOR,
                                 fg=TEXT_COLOR, font=("Arial", 10))
        generate_button.pack(pady=15, padx=10, fill=BOTH)

        # 创建右侧框架用于放置 Listbox
        right_frame = Frame(tab, bg=FRAME_BG_COLOR, bd=2, relief="groove")
        right_frame.pack(side=RIGHT, padx=10, pady=10, fill=BOTH, expand=True)

        # 创建 Listbox 用于展示 Excel 文件列表
        self.file_listbox = Listbox(right_frame, bg=BG_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        self.file_listbox.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)

        # 创建滚动条
        scrollbar = Scrollbar(right_frame, command=self.file_listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

    def select_files(self):
        # 打开文件选择对话框，允许多选 Excel 文件
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_paths:
            for path in file_paths:
                self.file_listbox.insert(END, path)
            self.file_paths_var.set(",".join(file_paths))
            self.status_label.config(text=f"已选择 {len(file_paths)} 个 Excel 文件")
        else:
            self.status_label.config(text="未选择任何 Excel 文件")

    def select_output_folder(self):
        # 打开文件夹选择对话框
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_var.set(folder_path)
            self.status_label.config(text=f"输出文件夹已选择: {folder_path}")
        else:
            self.status_label.config(text="未选择输出文件夹")

    def generate_pdf(self):
        folder_path = self.output_folder_var.get()
        if not folder_path:
            messagebox.showwarning("警告", "请先选择输出文件夹")
            return

        selected_paths = list(self.file_listbox.get(0, END))
        if not selected_paths:
            messagebox.showwarning("警告", "请先选择 Excel 文件")
            return

        import win32com.client
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False

        for excel_path in selected_paths:
            try:
                workbook = excel.Workbooks.Open(excel_path)
                pdf_filename = os.path.splitext(os.path.basename(excel_path))[0] + ".pdf"
                pdf_path = os.path.join(folder_path, pdf_filename)
                workbook.ExportAsFixedFormat(0, pdf_path)
                workbook.Close()
                messagebox.showinfo("成功", f"Excel 转 PDF 文件已生成: {pdf_path}")
                self.status_label.config(text=f"Excel 转 PDF 文件已生成: {pdf_path}")
            except Exception as e:
                messagebox.showerror("错误", f"生成 Excel PDF 文件时出错: {e}")
                logging.error(f"生成 Excel PDF 文件时出错: {e}")

        excel.Quit()


class OfficeToPDFConverter:
    def __init__(self, tab):
        self.tab = tab

        # 创建变量存储选择的文件路径和输出文件夹路径
        self.file_paths_var = StringVar()
        self.output_folder_var = StringVar()
        self.selected_option = StringVar()
        self.selected_option.set("word")  # 默认选择 Word 转 PDF

        # 创建标签
        self.status_label = Label(tab, text="请选择转换类型和文件", wraplength=380, bg=BG_COLOR, fg=TEXT_COLOR,
                                  font=("Arial", 12))
        self.status_label.pack(pady=10)

        # 创建 Radiobutton 选项
        Radiobutton(tab, text="Word转PDF", variable=self.selected_option, value="word", bg=BG_COLOR).pack()
        Radiobutton(tab, text="PPT转PDF", variable=self.selected_option, value="ppt", bg=BG_COLOR).pack()
        # 新增 Excel 转 PDF 选项
        Radiobutton(tab, text="Excel转PDF", variable=self.selected_option, value="excel", bg=BG_COLOR).pack()

        # 创建左侧框架用于放置按钮
        left_frame = Frame(tab, bg=FRAME_BG_COLOR, bd=2, relief="groove")
        left_frame.pack(side=LEFT, padx=10, pady=10, fill=Y)

        # 创建选择文件按钮
        self.select_file_button = Button(left_frame, text="选择文件", command=self.select_files, bg=BUTTON_COLOR,
                                         fg=TEXT_COLOR, font=("Arial", 10))
        self.select_file_button.pack(pady=5, padx=10, fill=BOTH)

        # 创建选择输出文件夹按钮
        select_folder_button = Button(left_frame, text="选择输出文件夹", command=self.select_output_folder,
                                      bg=BUTTON_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        select_folder_button.pack(pady=5, padx=10, fill=BOTH)

        # 创建生成 PDF 按钮
        generate_button = Button(left_frame, text="生成 PDF", command=self.generate_pdf, bg=BUTTON_COLOR,
                                 fg=TEXT_COLOR, font=("Arial", 10))
        generate_button.pack(pady=15, padx=10, fill=BOTH)

        # 创建右侧框架用于放置 Listbox
        right_frame = Frame(tab, bg=FRAME_BG_COLOR, bd=2, relief="groove")
        right_frame.pack(side=RIGHT, padx=10, pady=10, fill=BOTH, expand=True)

        # 创建 Listbox 用于展示文件列表
        self.file_listbox = Listbox(right_frame, bg=BG_COLOR, fg=TEXT_COLOR, font=("Arial", 10))
        self.file_listbox.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)

        # 创建滚动条
        scrollbar = Scrollbar(right_frame, command=self.file_listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

    def select_files(self):
        if self.selected_option.get() == "word":
            # 打开文件选择对话框，允许多选 Word 文件
            file_paths = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx;*.doc")])
        elif self.selected_option.get() == "ppt":
            # 打开文件选择对话框，允许多选 PPT 文件
            file_paths = filedialog.askopenfilenames(filetypes=[("PPT files", "*.ppt;*.pptx")])
        elif self.selected_option.get() == "excel":
            # 打开文件选择对话框，允许多选 Excel 文件
            file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_paths:
            for path in file_paths:
                self.file_listbox.insert(END, path)
            self.file_paths_var.set(",".join(file_paths))
            self.status_label.config(text=f"已选择 {len(file_paths)} 个文件")
        else:
            self.status_label.config(text="未选择任何文件")

    def select_output_folder(self):
        # 打开文件夹选择对话框
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_var.set(folder_path)
            self.status_label.config(text=f"输出文件夹已选择: {folder_path}")
        else:
            self.status_label.config(text="未选择输出文件夹")

    def generate_pdf(self):
        folder_path = self.output_folder_var.get()
        if not folder_path:
            messagebox.showwarning("警告", "请先选择输出文件夹")
            return

        selected_paths = list(self.file_listbox.get(0, END))
        if not selected_paths:
            messagebox.showwarning("警告", "请先选择文件")
            return

        if self.selected_option.get() == "word":
            for word_path in selected_paths:
                try:
                    pdf_filename = os.path.splitext(os.path.basename(word_path))[0] + ".pdf"
                    pdf_path = os.path.join(folder_path, pdf_filename)
                    docx2pdf.convert(word_path, pdf_path)
                    messagebox.showinfo("成功", f"Word 转 PDF 文件已生成: {pdf_path}")
                    self.status_label.config(text=f"Word 转 PDF 文件已生成: {pdf_path}")
                except Exception as e:
                    messagebox.showerror("错误", f"生成 Word PDF 文件时出错: {e}")
                    logging.error(f"生成 Word PDF 文件时出错: {e}")
        elif self.selected_option.get() == "ppt":
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            for ppt_path in selected_paths:
                try:
                    pdf_filename = os.path.splitext(os.path.basename(ppt_path))[0] + ".pdf"
                    pdf_path = os.path.join(folder_path, pdf_filename)
                    presentation = powerpoint.Presentations.Open(ppt_path)
                    presentation.SaveAs(pdf_path, 32)
                    presentation.Close()
                    messagebox.showinfo("成功", f"PPT 转 PDF 文件已生成: {pdf_path}")
                    self.status_label.config(text=f"PPT 转 PDF 文件已生成: {pdf_path}")
                except Exception as e:
                    messagebox.showerror("错误", f"生成 PPT PDF 文件时出错: {e}")
                    logging.error(f"生成 PPT PDF 文件时出错: {e}")
            powerpoint.Quit()
        elif self.selected_option.get() == "excel":
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            for excel_path in selected_paths:
                try:
                    workbook = excel.Workbooks.Open(excel_path)
                    pdf_filename = os.path.splitext(os.path.basename(excel_path))[0] + ".pdf"
                    pdf_path = os.path.join(folder_path, pdf_filename)
                    workbook.ExportAsFixedFormat(0, pdf_path)
                    workbook.Close()
                    messagebox.showinfo("成功", f"Excel 转 PDF 文件已生成: {pdf_path}")
                    self.status_label.config(text=f"Excel 转 PDF 文件已生成: {pdf_path}")
                except Exception as e:
                    messagebox.showerror("错误", f"生成 Excel PDF 文件时出错: {e}")
                    logging.error(f"生成 Excel PDF 文件时出错: {e}")
            excel.Quit()

    def select_files(self):
        if self.selected_option.get() == "word":
            # 打开文件选择对话框，允许多选 Word 文件
            file_paths = filedialog.askopenfilenames(filetypes=[("Word files", "*.docx;*.doc")])
        elif self.selected_option.get() == "ppt":
            # 打开文件选择对话框，允许多选 PPT 文件
            file_paths = filedialog.askopenfilenames(filetypes=[("PPT files", "*.ppt;*.pptx")])
        elif self.selected_option.get() == "excel":
            # 打开文件选择对话框，允许多选 Excel 文件
            file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_paths:
            for path in file_paths:
                self.file_listbox.insert(END, path)
            self.file_paths_var.set(",".join(file_paths))
            self.status_label.config(text=f"已选择 {len(file_paths)} 个文件")
        else:
            self.status_label.config(text="未选择任何文件")

    def generate_pdf(self):
        folder_path = self.output_folder_var.get()
        if not folder_path:
            messagebox.showwarning("警告", "请先选择输出文件夹")
            return

        selected_paths = list(self.file_listbox.get(0, END))
        if not selected_paths:
            messagebox.showwarning("警告", "请先选择文件")
            return

        if self.selected_option.get() == "word":
            for word_path in selected_paths:
                try:
                    pdf_filename = os.path.splitext(os.path.basename(word_path))[0] + ".pdf"
                    pdf_path = os.path.join(folder_path, pdf_filename)
                    docx2pdf.convert(word_path, pdf_path)
                    messagebox.showinfo("成功", f"Word 转 PDF 文件已生成: {pdf_path}")
                    self.status_label.config(text=f"Word 转 PDF 文件已生成: {pdf_path}")
                except Exception as e:
                    messagebox.showerror("错误", f"生成 Word PDF 文件时出错: {e}")
                    logging.error(f"生成 Word PDF 文件时出错: {e}")
        elif self.selected_option.get() == "ppt":
            powerpoint = win32com.client.Dispatch("PowerPoint.Application")
            for ppt_path in selected_paths:
                try:
                    pdf_filename = os.path.splitext(os.path.basename(ppt_path))[0] + ".pdf"
                    pdf_path = os.path.join(folder_path, pdf_filename)
                    presentation = powerpoint.Presentations.Open(ppt_path)
                    presentation.SaveAs(pdf_path, 32)
                    presentation.Close()
                    messagebox.showinfo("成功", f"PPT 转 PDF 文件已生成: {pdf_path}")
                    self.status_label.config(text=f"PPT 转 PDF 文件已生成: {pdf_path}")
                except Exception as e:
                    messagebox.showerror("错误", f"生成 PPT PDF 文件时出错: {e}")
                    logging.error(f"生成 PPT PDF 文件时出错: {e}")
            powerpoint.Quit()
        elif self.selected_option.get() == "excel":
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            for excel_path in selected_paths:
                try:
                    workbook = excel.Workbooks.Open(excel_path)
                    pdf_filename = os.path.splitext(os.path.basename(excel_path))[0] + ".pdf"
                    pdf_path = os.path.join(folder_path, pdf_filename)
                    workbook.ExportAsFixedFormat(0, pdf_path)
                    workbook.Close()
                    messagebox.showinfo("成功", f"Excel 转 PDF 文件已生成: {pdf_path}")
                    self.status_label.config(text=f"Excel 转 PDF 文件已生成: {pdf_path}")
                except Exception as e:
                    messagebox.showerror("错误", f"生成 Excel PDF 文件时出错: {e}")
                    logging.error(f"生成 Excel PDF 文件时出错: {e}")
            excel.Quit()

def show_about():
    about_window = Toplevel(root)
    about_window.title("关于")
    about_window_width = 300
    about_window_height = 150
    about_window.geometry(f"{about_window_width}x{about_window_height}")
    about_window.configure(bg=BG_COLOR)

    # 获取主窗口的位置和大小
    root.update_idletasks()
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    root_width = root.winfo_width()
    root_height = root.winfo_height()

    # 计算关于窗口的位置
    about_x = root_x + (root_width - about_window_width) // 2
    about_y = root_y + (root_height - about_window_height) // 2

    # 设置关于窗口的位置
    about_window.geometry(f"+{about_x}+{about_y}")

    author_label = Label(about_window, text="作者：牛逼神仙", bg=BG_COLOR, fg=TEXT_COLOR, font=("Arial", 12))
    author_label.pack(pady=20)

    email_label = Label(about_window, text="电子邮件：linshenyang@qq.com", bg=BG_COLOR, fg=TEXT_COLOR, font=("Arial", 12))
    email_label.pack(pady=10)


if __name__ == "__main__":
    root = Tk()
    root.title("文件转换工具 - 作者：牛逼神仙")
    root.geometry("800x600")
    root.configure(bg=BG_COLOR)

    # 获取屏幕宽度和高度
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 获取窗口宽度和高度
    window_width = 800
    window_height = 600

    # 计算窗口在屏幕上居中的坐标
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2

    # 设置窗口的几何位置，使其居中
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")


    # 创建菜单栏
    menubar = Menu(root)
    about_menu = Menu(menubar, tearoff=0)
    about_menu.add_command(label="关于", command=show_about)
    menubar.add_cascade(label="关于", menu=about_menu)
    root.config(menu=menubar)

    # 创建 Notebook 控件（TabControl）
    notebook = Notebook(root)
    notebook.pack(fill=BOTH, expand=True)

    # 创建图片转 PDF 标签页
    pdf_tab = Frame(notebook, bg=BG_COLOR)
    notebook.add(pdf_tab, text="图片转 PDF")
    ImageToPDFConverter(pdf_tab)

    # 创建 Office 转 PDF 标签页
    office_tab = Frame(notebook, bg=BG_COLOR)
    notebook.add(office_tab, text="Office文档转换")
    OfficeToPDFConverter(office_tab)

    # 创建 PDF 转其他格式标签页
    pdf_to_file_tab = Frame(notebook, bg="#f0f0f0")
    notebook.add(pdf_to_file_tab, text="PDF转换")
    PDFToFileConverter(pdf_to_file_tab)

    root.mainloop()