import fitz  # PyMuPDF
from PIL import Image
import io
import tkinter as tk
from tkinter import filedialog
from functools import partial


def get_num(total_page, n):
    if n % 2 == 1:
        y = int((n + 1) / 2)
        return total_page - 2 * (y - 1), 2 * (y - 1) + 1
    else:
        y = int(n / 2)
        return 2 * y, total_page - 1 - 2 * (y - 1)


class PDF:
    def __init__(self):
        self.doc = fitz.open()

    def add_page(self, img1, img2, is_c):
        a4_width = 842
        a4_height = 595
        a5_width = 420
        a5_height = 595
        # A4、A5纸尺寸

        page = self.doc.new_page(width=a4_width, height=a4_height)
        # 创建新界面

        img1_buffer = io.BytesIO()
        img1.save(img1_buffer, format='PNG')
        img1_data = img1_buffer.getvalue()
        img1_buffer.close()

        left_rect = fitz.Rect(0, 0, a5_width, a5_height)
        page.insert_image(left_rect, stream=img1_data)
        # 放入左侧图片

        img2_buffer = io.BytesIO()
        img2.save(img2_buffer, format='PNG')
        img2_data = img2_buffer.getvalue()
        img2_buffer.close()

        right_rect = fitz.Rect(a5_width, 0, a4_width, a5_height)
        page.insert_image(right_rect, stream=img2_data)
        # 放入右侧图片

        img1.close()
        img2.close()
        # 释放内存

        if is_c:
            page.set_rotation(180)

    def save(self, pdf_path):
        self.doc.save(pdf_path)
        self.doc.close()


class RawPDF:
    def __init__(self, pdf_path):
        self.doc = fitz.open(pdf_path)
        self.total_pages = len(self.doc)

    def get_total_pages(self):  # 返回pdf页数
        return self.total_pages

    def pdf_page_to_image(self, n, zoom=2.5):  # 将第n页转为图片
        if n > self.total_pages:
            img = Image.new('RGB', (420, 595), 'white')
            # print("超出文件页数")
        else:
            mat = fitz.Matrix(zoom, zoom)
            pix = self.doc.load_page(n - 1).get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            img = Image.open(io.BytesIO(img_data))
        return img

    def close(self):  # 关闭文档
        self.doc.close()


def run(raw_path, is_change):
    path = raw_path[:-4]+"(小册子).pdf"
    raw_doc = RawPDF(raw_path)
    # 打开源文件

    raw_page = raw_doc.get_total_pages()
    if raw_page % 4 == 0:
        pdf_page = int(raw_page / 4)
    else:
        pdf_page = int(raw_page // 4 + 1)
    # 获取源文件页数并计算需要的纸张数量(文件页数应是其两倍）

    pdf = PDF()
    # 创建新文件

    for i in range(pdf_page * 2):
        page1_num, page2_num = get_num(pdf_page * 4, i + 1)
        page1 = raw_doc.pdf_page_to_image(page1_num)
        page2 = raw_doc.pdf_page_to_image(page2_num)
        pdf.add_page(page1, page2, (is_change.get() and (i + 1) % 2 == 0))
    # 加入新页面

    pdf.save(path)
    # 保存文件

    lab = tk.Label(root, text="完成！已保存至同文件夹", font=("Arial", 24))
    lab.pack(pady=5)  # pady 设置垂直间距


root = tk.Tk()
root.title("pdf打印小工具")
root.geometry("600x370")

file_path = filedialog.askopenfilename(
    title="选择文件",
    filetypes=[("PDF文件", "*.pdf")]
)

label = tk.Label(root, text="已选择:", font=("Arial", 24))
label.pack(pady=5)  # pady 设置垂直间距
label = tk.Label(root, text=str(file_path).split("/")[-1], font=("Arial", 24))
label.pack(pady=5)
# 确认已选择的文档

is_checked = tk.BooleanVar(value=True)
checkbox = tk.Checkbutton(
    root,
    text="需要翻转偶数面",
    font=("Arial", 24),
    variable=is_checked
)
checkbox.pack(pady=5)
# 放置复选框

label = tk.Label(root, text="请确保是双面打印", font=("Arial", 24))
label.pack(pady=5)
# 提示

button = tk.Button(root, text="运行", font=("Arial", 24), command=partial(run, file_path, is_checked))
button.pack(pady=5)

root.mainloop()
