import os
from pathlib import Path
import requests
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from datetime import datetime
import re
import urllib3
from bs4 import BeautifulSoup
import time
from tkinter import Tk, Label, Button, filedialog, messagebox
import tkinter as tk
import webbrowser

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Load the KML file
# kml_file_path = '11.kml'


# 转换十进制度数为度分秒格式
def decimal_to_dms(value, is_latitude=True):
    degrees = int(value)
    minutes = int((abs(value) - abs(degrees)) * 60)
    seconds = (abs(value) - abs(degrees) - minutes / 60) * 3600
    direction = 'N' if is_latitude and value >= 0 else 'S' if is_latitude else 'E' if value >= 0 else 'W'
    return f"{abs(degrees)}°{minutes}'{seconds:.2f}″{direction}"


# 下载图片并保存，带有重试机制
def download_image(url, name, save_dir, retries=3):
    try:
        for attempt in range(retries):
            try:
                response = requests.get(url, stream=True, verify=False)
                if response.status_code == 200:
                    # 使用 name_text 命名文件
                    safe_name = re.sub(r'[\\/*?:"<>|]', "_", name)  # 替换不安全字符
                    filename = f"{safe_name}.jpg"  # 假设你希望使用 .jpg 后缀
                    file_path = save_dir / filename
                    with open(file_path, 'wb') as f:
                        for chunk in response.iter_content(1024):
                            f.write(chunk)
                    return file_path.relative_to(save_dir.parent)  # 返回相对路径
            except requests.RequestException as e:
                print(f"尝试 {attempt + 1}/{retries}：图片下载失败：{e}")
                time.sleep(1)  # 等待一秒后重试
        return ""  # 所有尝试失败后返回空字符串
    except Exception as e:
        print(f"意外错误：{e}")
    return ""


def kml_to_xlsx(kml_file, export_dir):
    # 获取所在文件夹
    input_path = Path(kml_file)
    photos_dir = input_path.parent / "photos"

    # 新建照片文件夹（如果不存在）
    os.makedirs(photos_dir, exist_ok=True)

    # 解析 KML 文件
    tree = ET.parse(kml_file)
    root = tree.getroot()

    # 创建一个新的 Excel 工作簿
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "KML Data"

    # 设置标题行
    headers = ["采集号", "物种中文", "拉丁名", "所属岛屿", "生境", "东经", "北纬", "海拔", "图片", "备注"]
    sheet.append(headers)

    # 定义命名空间
    namespace = {'kml': 'http://www.opengis.net/kml/2.2'}


    # 遍历每个 Placemark 元素
    for placemark in root.findall(".//kml:Placemark", namespace):
        if placemark.get("id") != "realPoint":
            continue

        # Extract name, description, and date (if available)
        name = placemark.find("kml:name", namespace)
        description = placemark.find("kml:description", namespace)
        timestamp = placemark.find(".//kml:TimeStamp/kml:when", namespace)
        coordinates = placemark.find(".//kml:coordinates", namespace)

        # 这里初始化 photo 变量
        photo = ""
        desc_words = ""
        island=""

        if description is not None and description.text:
            # 使用 BeautifulSoup 解析描述
            soup = BeautifulSoup(description.text, 'html.parser')
            img_tag = soup.find('img')  # 查找 <img> 标签
            if img_tag and img_tag.has_attr('src'):
                img_url = img_tag['src']
                img_filename = img_url.split("/")[-1].split("?")[0]  # 提取文件名
                if name is not None and name.text is not None:
                    name_text = name.text.strip()  # 确保 name_text 不为空
                    # img_path = download_image(img_url, name_text, photos_dir)
                    # photo = str(img_path) if img_path else ""

                # 提取岛屿信息（从文件名中提取）
                # Fix
                island = img_filename.split('?')[0]  # 从图片 URL 中提取文件名作为岛屿名称

            # 提取描述文本
            a_tag = soup.find('a')
            if a_tag:
                # 提取 <a> 标签中的文本，不包括 <img> 的内容
                desc_words = a_tag.get_text(strip=True)

        # 获取名称、描述和日期（如果存在）
        name_text = name.text.strip() if name is not None and name.text is not None else ""
        date_text = ""
        if timestamp is not None and timestamp.text is not None:
            original_date = timestamp.text
            date_text = datetime.strptime(original_date, "%Y-%m-%dT%H:%M:%SZ").strftime("%Y%m%d")
        coord_text = coordinates.text.strip() if coordinates is not None else ""

        name_text = date_text + name_text

        # 提取经度和纬度（忽略高度）
        if coord_text:
            lon, lat, hig = coord_text.split(",")
            lat_dms = decimal_to_dms(float(lat), is_latitude=True)
            lon_dms = decimal_to_dms(float(lon), is_latitude=False)
            # 创建超链接，格式为：('显示文本', '链接')
            if photo:
                photo_link = f'=HYPERLINK("{photo}", "查看图片")'
            else:
                photo_link = ""  # 如果没有照片，留空
            row = [name_text, "", "", "", "", lon_dms, lat_dms, float(hig), photo_link, desc_words]
        else:
            row = [name_text, "", "", "d1", "", "", "", "", "pic", ""]

        # 将数据添加到 Excel 表中
        sheet.append(row)
    # # 保存工作簿
    # workbook.save(xlsx_file)
    # return xlsx_file

    # 保存 Excel 文件
    export_path = export_dir / f"{Path(kml_file).stem}_output.xlsx"
    workbook.save(export_path)
    messagebox.showinfo("完成", f"文件已成功导出至: {export_path}")
    return export_path


# GUI
def select_kml_file():
    file_path = filedialog.askopenfilename(filetypes=[("KML files", "*.kml")])
    if file_path:
        kml_file_var.set(file_path)
        kml_file_label.config(text=f"已选择文件: {file_path}")

        dir_path=os.path.dirname(file_path)
        export_dir_var.set(dir_path)
        export_dir_label.config(text=f"导出目录: {dir_path}")


def select_export_dir():
    dir_path = filedialog.askdirectory()
    if dir_path:
        export_dir_var.set(dir_path)
        export_dir_label.config(text=f"导出目录: {dir_path}")

def generate_excel():
    kml_file = kml_file_var.get()
    export_dir = Path(export_dir_var.get())
    if kml_file and export_dir:
        export_path = kml_to_xlsx(kml_file, export_dir)
        open_file_button.config(command=lambda: webbrowser.open(export_path), state="normal")
    else:
        messagebox.showwarning("警告", "请先选择KML文件和导出目录")

# 主窗口设置
root = Tk()
root.title("KML转Excel工具")
root.geometry("400x400")

# 设置界面元素
kml_file_var = tk.StringVar()
export_dir_var = tk.StringVar()

Label(root, text="选择KML文件：").pack(pady=5)
Button(root, text="选择文件", command=select_kml_file).pack(pady=5)
kml_file_label = Label(root, text="")
kml_file_label.pack()

Label(root, text="选择导出目录：").pack(pady=5)
Button(root, text="选择目录", command=select_export_dir).pack(pady=5)
export_dir_label = Label(root, text="")
export_dir_label.pack()

Button(root, text="生成Excel文件", command=generate_excel, bg="green", fg="white").pack(pady=10)

# 添加打开文件按钮
open_file_button = Button(root, text="打开生成的文件", state="disabled")
open_file_button.pack(pady=5)

# 启动主循环
root.mainloop()

# # 定义输出文件
# xlsx_output_path = '11out.xlsx'
# kml_to_xlsx(kml_file_path, xlsx_output_path)
# print(f"输出文件已保存至：{xlsx_output_path}")


