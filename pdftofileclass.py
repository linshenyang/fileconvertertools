import os
import logging
from tkinter import Button, Label, filedialog, messagebox, StringVar, Radiobutton, Listbox, Scrollbar, Frame, LEFT, RIGHT, Y, BOTH, END, Toplevel
from tkinter.ttk import Notebook
from PIL import Image
import fitz  # PyMuPDF
import docx
import pandas as pd
import win32com.client  # 用于 PPT 转 PDF
from bs4 import BeautifulSoup
import io

# 配置日志记录
logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')

class PDFToFileConverter:
    def __init__(self, tab):
        self.tab = tab

        # 创建变量存储选择的 PDF 文件路径和输出文件夹路径
        self.file_paths_var = StringVar()
        self.output_folder_var = StringVar()
        self.selected_option = StringVar()
        self.selected_option.set("word")  # 默认选择 PDF 转 Word

        # 创建标签
        self.status_label = Label(tab, text="请选择转换类型和文件", wraplength=380, bg="#f0f0f0", fg="#333333",
                                  font=("Arial", 12))
        self.status_label.pack(pady=10)

        # 创建 Radiobutton 选项
        Radiobutton(tab, text="PDF转Word", variable=self.selected_option, value="word", bg="#f0f0f0").pack()
        Radiobutton(tab, text="PDF转图片", variable=self.selected_option, value="image", bg="#f0f0f0").pack()
        Radiobutton(tab, text="PDF转Excel", variable=self.selected_option, value="excel", bg="#f0f0f0").pack()
        Radiobutton(tab, text="PDF转PPT", variable=self.selected_option, value="ppt", bg="#f0f0f0").pack()
        Radiobutton(tab, text="PDF转TXT", variable=self.selected_option, value="txt", bg="#f0f0f0").pack()
        Radiobutton(tab, text="PDF转HTML", variable=self.selected_option, value="html", bg="#f0f0f0").pack()
        Radiobutton(tab, text="PDF转长图", variable=self.selected_option, value="long_image", bg="#f0f0f0").pack()

        # 创建左侧框架用于放置按钮
        left_frame = Frame(tab, bg="#ffffff", bd=2, relief="groove")
        left_frame.pack(side=LEFT, padx=10, pady=10, fill=Y)

        # 创建选择文件按钮
        self.select_file_button = Button(left_frame, text="选择文件", command=self.select_files, bg="#e0e0e0",
                                         fg="#333333", font=("Arial", 10))
        self.select_file_button.pack(pady=5, padx=10, fill=BOTH)

        # 创建选择输出文件夹按钮
        select_folder_button = Button(left_frame, text="选择输出文件夹", command=self.select_output_folder,
                                      bg="#e0e0e0", fg="#333333", font=("Arial", 10))
        select_folder_button.pack(pady=5, padx=10, fill=BOTH)

        # 创建生成文件按钮
        generate_button = Button(left_frame, text="生成文件", command=self.generate_file, bg="#e0e0e0",
                                 fg="#333333", font=("Arial", 10))
        generate_button.pack(pady=15, padx=10, fill=BOTH)

        # 创建右侧框架用于放置 Listbox
        right_frame = Frame(tab, bg="#ffffff", bd=2, relief="groove")
        right_frame.pack(side=RIGHT, padx=10, pady=10, fill=BOTH, expand=True)

        # 创建 Listbox 用于展示文件列表
        self.file_listbox = Listbox(right_frame, bg="#f0f0f0", fg="#333333", font=("Arial", 10))
        self.file_listbox.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)

        # 创建滚动条
        scrollbar = Scrollbar(right_frame, command=self.file_listbox.yview)
        scrollbar.pack(side=RIGHT, fill=Y)
        self.file_listbox.config(yscrollcommand=scrollbar.set)

    def select_files(self):
        # 打开文件选择对话框，允许多选 PDF 文件
        file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if file_paths:
            for path in file_paths:
                self.file_listbox.insert(END, path)
            self.file_paths_var.set(",".join(file_paths))
            self.status_label.config(text=f"已选择 {len(file_paths)} 个 PDF 文件")
        else:
            self.status_label.config(text="未选择任何 PDF 文件")

    def select_output_folder(self):
        # 打开文件夹选择对话框
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.output_folder_var.set(folder_path)
            self.status_label.config(text=f"输出文件夹已选择: {folder_path}")
        else:
            self.status_label.config(text="未选择输出文件夹")

    def generate_file(self):
        folder_path = self.output_folder_var.get()
        if not folder_path:
            messagebox.showwarning("警告", "请先选择输出文件夹")
            return

        selected_paths = list(self.file_listbox.get(0, END))
        if not selected_paths:
            messagebox.showwarning("警告", "请先选择 PDF 文件")
            return

        for pdf_path in selected_paths:
            try:
                if self.selected_option.get() == "word":
                    self.pdf_to_word(pdf_path, folder_path)
                elif self.selected_option.get() == "image":
                    self.pdf_to_images(pdf_path, folder_path)
                elif self.selected_option.get() == "excel":
                    self.pdf_to_excel(pdf_path, folder_path)
                elif self.selected_option.get() == "ppt":
                    self.pdf_to_ppt(pdf_path, folder_path)
                elif self.selected_option.get() == "txt":
                    self.pdf_to_txt(pdf_path, folder_path)
                elif self.selected_option.get() == "html":
                    self.pdf_to_html(pdf_path, folder_path)
                elif self.selected_option.get() == "long_image":
                    self.pdf_to_long_image(pdf_path, folder_path)
            except Exception as e:
                messagebox.showerror("错误", f"转换文件时出错: {e}")
                logging.error(f"转换文件时出错: {e}")

    def pdf_to_word(self, pdf_path, folder_path):
        doc = docx.Document()
        pdf_document = fitz.open(pdf_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text("text")
            doc.add_paragraph(text)
        output_path = os.path.join(folder_path, os.path.splitext(os.path.basename(pdf_path))[0] + ".docx")
        doc.save(output_path)
        messagebox.showinfo("成功", f"PDF 转 Word 文件已生成: {output_path}")
        self.status_label.config(text=f"PDF 转 Word 文件已生成: {output_path}")

    def pdf_to_images(self, pdf_path, folder_path):
        pdf_document = fitz.open(pdf_path)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()
            output_path = os.path.join(folder_path, f"{os.path.splitext(os.path.basename(pdf_path))[0]}_page_{page_num + 1}.png")
            pix.save(output_path)
        messagebox.showinfo("成功", f"PDF 转图片文件已生成: {folder_path}")
        self.status_label.config(text=f"PDF 转图片文件已生成: {folder_path}")

    def pdf_to_excel(self, pdf_path, folder_path):
        pdf_document = fitz.open(pdf_path)
        workbook = pd.ExcelWriter(os.path.join(folder_path, os.path.splitext(os.path.basename(pdf_path))[0] + ".xlsx"), engine='xlsxwriter')
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            text = page.get_text("text")
            df = pd.DataFrame([text.split('\n')])
            df.to_excel(workbook, sheet_name=f'Page {page_num + 1}', index=False, header=False)
        workbook.save()
        messagebox.showinfo("成功", f"PDF 转 Excel 文件已生成: {folder_path}")
        self.status_label.config(text=f"PDF 转 Excel 文件已生成: {folder_path}")

    def pdf_to_ppt(self, pdf_path, folder_path):
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Add()
        pdf_document = fitz.open(pdf_path)
        for page_num in range(len(pdf_document)):
            slide = presentation.Slides.Add(page_num + 1, 12)  # 12 is ppLayoutText
            page = pdf_document.load_page(page_num)
            text = page.get_text("text")
            slide.Shapes.Title.TextFrame.TextRange.Text = f"Page {page_num + 1}"
            slide.Shapes.Placeholders(2).TextFrame.TextRange.Text = text
        output_path = os.path.join(folder_path, os.path.splitext(os.path.basename(pdf_path))[0] + ".pptx")
        presentation.SaveAs(output_path)
        presentation.Close()
        powerpoint.Quit()
        messagebox.showinfo("成功", f"PDF 转 PPT 文件已生成: {output_path}")
        self.status_label.config(text=f"PDF 转 PPT 文件已生成: {output_path}")

    def pdf_to_txt(self, pdf_path, folder_path):
        pdf_document = fitz.open(pdf_path)
        with open(os.path.join(folder_path, os.path.splitext(os.path.basename(pdf_path))[0] + ".txt"), "w", encoding="utf-8") as txt_file:
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                text = page.get_text("text")
                txt_file.write(text)
        messagebox.showinfo("成功", f"PDF 转 TXT 文件已生成: {folder_path}")
        self.status_label.config(text=f"PDF 转 TXT 文件已生成: {folder_path}")

    def pdf_to_html(self, pdf_path, folder_path):
        pdf_document = fitz.open(pdf_path)
        with open(os.path.join(folder_path, os.path.splitext(os.path.basename(pdf_path))[0] + ".html"), "w", encoding="utf-8") as html_file:
            html_file.write("<html><body>")
            for page_num in range(len(pdf_document)):
                page = pdf_document.load_page(page_num)
                text = page.get_text("html")
                soup = BeautifulSoup(text, 'html.parser')
                html_file.write(str(soup))
            html_file.write("</body></html>")
        messagebox.showinfo("成功", f"PDF 转 HTML 文件已生成: {folder_path}")
        self.status_label.config(text=f"PDF 转 HTML 文件已生成: {folder_path}")

    def pdf_to_long_image(self, pdf_path, folder_path):
        pdf_document = fitz.open(pdf_path)
        images = []
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap()
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(img)
        long_image = Image.new('RGB', (images[0].width, sum(img.height for img in images)))
        y_offset = 0
        for img in images:
            long_image.paste(img, (0, y_offset))
            y_offset += img.height
        output_path = os.path.join(folder_path, os.path.splitext(os.path.basename(pdf_path))[0] + "_long_image.png")
        long_image.save(output_path)
        messagebox.showinfo("成功", f"PDF 转长图文件已生成: {output_path}")
        self.status_label.config(text=f"PDF 转长图文件已生成: {output_path}")