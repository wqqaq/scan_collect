import requests
from bs4 import BeautifulSoup
import openpyxl

def extract_fields_from_url(url):
    # 发送HTTP请求获取页面内容
    response = requests.get(url)

    # 检查响应状态码是否为200
    if response.status_code == 200:
        html_content = response.text

        # 使用BeautifulSoup解析HTML
        soup = BeautifulSoup(html_content, 'html.parser')

        # 提取所有段落文本并查找字段
        paragraphs = soup.find_all('p')
        found_fields = False  # 标记是否找到字段

        severity = "No known"
        issue_type = "No known"
        exploit_available = "No known"

        for p in paragraphs:
            text = p.text.strip()
            if "Severity" in text:
                severity = text.split(":")[1].strip()  # 提取严重程度字段值
                found_fields = True
            elif "Type" in text:
                issue_type = text.split(":")[1].strip()  # 提取类型字段值
                found_fields = True
            elif "Exploit Available" in text:
                exploit_available = text.split(":")[1].strip()  # 提取是否有利用程序可用字段值
                found_fields = True

        if not found_fields:
            return None

        return severity, issue_type, exploit_available
    return None

# 创建一个新的Excel工作簿，并写入标题行
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet['A1'] = "URL"
sheet['B1'] = "Severity"
sheet['C1'] = "Type"
sheet['D1'] = "Exploit Available"

# 处理每个URL并写入Excel
urls = []

with open("wq.txt", "r") as file:
    for line in file:
        url = "https://www.tenable.com/plugins/nessus/" + line
        fields = extract_fields_from_url(url)
        
        if fields is not None:
            row = [url] + list(fields)
            sheet.append(row)

# 保存Excel文件
workbook.save('nessus_plugin.xlsx')

print("所有URL处理完毕")