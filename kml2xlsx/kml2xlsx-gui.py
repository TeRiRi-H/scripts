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
from concurrent.futures import ThreadPoolExecutor

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# 转换十进制度数为度分秒格式
def decimal_to_dms(value, is_latitude=True):
    degrees = int(value)
    minutes = int((abs(value) - abs(degrees)) * 60)
    seconds = (abs(value) - abs(degrees) - minutes / 60) * 3600
    direction = 'N' if is_latitude and value >= 0 else 'S' if is_latitude else 'E' if value >= 0 else 'W'
    return f"{abs(degrees)}°{minutes}'{seconds:.2f}″{direction}"


# 多线程下载图片
def download_image(url, name, save_dir, retries=3):
    try:
        for attempt in range(retries):
            try:
                response = requests.get(url, stream=True, verify=False)
                if response.status_code == 200:
                    safe_name = re.sub(r'[\\/*?:"<>|]', "_", name)
                    filename = f"{safe_name}.jpg"
                    file_path = save_dir / filename
                    with open(file_path, 'wb') as f:
                        for chunk in response.iter_content(1024):
                            f.write(chunk)
                    return file_path.relative_to(save_dir.parent)  # 返回相对路径
            except requests.RequestException as e:
                print(f"尝试 {attempt + 1}/{retries}：图片下载失败：{e}")
                time.sleep(1)  # 重试前等待
        return ""  # 若重试失败，返回空字符串
    except Exception as e:
        print(f"意外错误：{e}")
    return ""

def download_images_multithread(image_urls, names, save_dir):
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = [
            executor.submit(download_image, url, name, save_dir)
            for url, name in zip(image_urls, names)
        ]
        results = [future.result() for future in futures]
    return results


def kml_to_xlsx(kml_file, export_dir):
    input_path = Path(kml_file)
    photos_dir = input_path.parent / "photos"
    os.makedirs(photos_dir, exist_ok=True)

    tree = ET.parse(kml_file)
    root = tree.getroot()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "KML Data"
    headers = ["采集号", "物种中文", "拉丁名", "所属岛屿", "生境", "东经", "北纬", "海拔", "图片", "备注"]
    sheet.append(headers)

    namespace = {'kml': 'http://www.opengis.net/kml/2.2'}

    # 初始化图像 URL 和名称列表
    image_urls = []
    names = []
    rows_data = []  # 临时存储表格行数据

    for placemark in root.findall(".//kml:Placemark", namespace):
        if placemark.get("id") != "realPoint":
            continue

        name = placemark.find("kml:name", namespace)
        description = placemark.find("kml:description", namespace)
        timestamp = placemark.find(".//kml:TimeStamp/kml:when", namespace)
        coordinates = placemark.find(".//kml:coordinates", namespace)

        photo = ""
        desc_words = ""
        island = ""

        if description is not None and description.text:
            soup = BeautifulSoup(description.text, 'html.parser')
            img_tag = soup.find('img')
            if img_tag and img_tag.has_attr('src'):
                img_url = img_tag['src']
                img_filename = img_url.split("/")[-1].split("?")[0]
                if name is not None and name.text is not None:
                    name_text = name.text.strip()
                    image_urls.append(img_url)  # 收集 URL 和名称
                    names.append(name_text)
                island = img_filename.split('?')[0]

            a_tag = soup.find('a')
            if a_tag:
                desc_words = a_tag.get_text(strip=True)

        name_text = name.text.strip() if name is not None and name.text is not None else ""
        date_text = ""
        if timestamp is not None and timestamp.text is not None:
            original_date = timestamp.text
            date_text = datetime.strptime(original_date, "%Y-%m-%dT%H:%M:%SZ").strftime("%Y%m%d")
        coord_text = coordinates.text.strip() if coordinates is not None else ""

        name_text = date_text + name_text

        if coord_text:
            lon, lat, hig = coord_text.split(",")
            lat_dms = decimal_to_dms(float(lat), is_latitude=True)
            lon_dms = decimal_to_dms(float(lon), is_latitude=False)
            rows_data.append([name_text, "", "", "", "", lon_dms, lat_dms, float(hig), "", desc_words])
        else:
            rows_data.append([name_text, "", "", "", "", "", "", "", "", ""])

    # 执行图片的多线程下载
    image_paths = download_images_multithread(image_urls, names, photos_dir)

    # 将下载的图片路径添加到相应行的“图片”列
    for row, img_path in zip(rows_data, image_paths):
        row[8] = f'=HYPERLINK("{img_path}", "查看图片")' if img_path else ""
        sheet.append(row)  # 添加到 Excel 中

    export_path = export_dir / f"{input_path.stem}_output.xlsx"
    workbook.save(export_path)
    print(f"文件已成功导出至: {export_path}")
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


