import requests
from bs4 import BeautifulSoup
import pandas as pd
import openpyxl
import re
from datetime import datetime

# 初始化，Session来持久化
session = requests.Session()
base_url = 'https://www.kylinos.cn/'
url_template = 'https://www.kylinos.cn/support/loophole/patch.html?page={}'
F = '修复的CVE'
G = '受影响的软件包'
H = '软件包修复版本'
I = '修复方法'
J = '软件包下载地址'
K = '修复验证'
split_patterns = {
    F: (F,G),
    G: (G,H),
    H: (H,I),
    I: (I,J),
    J: (J,K),
    K: (K,None)
}
# 爬取字段
def extract_fields_from_soup(soup):
    fields_list = []
    # 查找所有符合条件的<a>标签
    for a_tag in soup.find_all('a', href=True, style=lambda x: x and 'color:#de0515;' in x):
        # 提取公告 ID
        announcement_id = a_tag.text.strip()
        # 查找包含描述和发布时间的父元素
        parent = a_tag.find_parent('tr')
        if parent:
            target_td = parent.find('td', style='color:#1f1f1f;')
            if target_td:
                # 在找到的 <td> 标签中查找 <span> 标签
                span = target_td.find('span')
                if span:
                    severity = span.text.strip()
                else:
                    severity = '未知'
            else:
                severity = '未知'
            description_tag = parent.find('td', class_='mobile-hide')
            description = description_tag.text.strip() if description_tag else '未知'
            # 提取发布时间
            release_date_tag = parent.find_all('td', class_='mobile-hide')[1]  # 假设发布时间是第二个
            release_date = release_date_tag.text.strip() if release_date_tag else '未知'
            # 提取详细介绍
            content = a_tag.get('href')  # 假设内容是链接的href属性

            # 构建完整的URL
            full_url = base_url + content

            # 发送HTTP请求获取公告页面内容
            try:
                announcement_response = session.get(full_url, timeout=10)
                if announcement_response.status_code == 200:
                    announcement_soup = BeautifulSoup(announcement_response.text, 'html.parser')
                    base_desc_divs = announcement_soup.find_all('div', class_='base-desc')
                    if base_desc_divs and len(base_desc_divs) > 1:
                        second_base_desc_div = base_desc_divs[1]
                        content = second_base_desc_div.get_text(strip=True)
            except requests.exceptions.RequestException as e:
                print(f"请求失败: {e}")
                continue

            # 将所有字段添加到fields_list列表
            fields_list.append([announcement_id, severity, description, release_date, content])

    return fields_list

# 清洗详细介绍列字段
def clean_content(content):
    content = content.replace('\r\n', '\n')
    content = content.replace('·', '')
    content = content.replace('受影响的操作系统及软件包', '受影响的软件包')
    content = content.replace('4.影响系统情况', '2.受影响的软件包')
    content = content.replace('2. 相关版本/架构', '')
    content = content.replace('3. 描述', '')
    content = content.replace('影响的操作系统及修复版本', '受影响的软件包')
    # 去除特定字符串前的空格、小标题
    titles = [F,G,H,I,J,K]
    pattern = r'(?:\d+\.\s*)?({})\s*'.format('|'.join(titles))
    # 使用正则表达式替换匹配到的空格
    clean_content = re.sub(pattern, r'\1 ', content, flags=re.IGNORECASE)
    return clean_content.strip()  # 去除字符串两端的空格

# 主函数
def main():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet['A1'] = "公告 ID"
    sheet['B1'] = "安全级别"
    sheet['C1'] = "描述"
    sheet['D1'] = "发布时间"
    sheet['E1'] = "详细介绍"

    df = pd.DataFrame(columns=['公告 ID', '安全级别', '描述', '发布时间', '详细介绍'])

    # 循环处理每个页面
    for page_number in range(1, 2):  # html的page范围
        url = url_template.format(page_number)
        try:
            response = session.get(url)
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, 'html.parser')
                fields = extract_fields_from_soup(soup)
                if fields:  # 确保fields不是None
                    for field in fields:
                        sheet.append(field)  # 将每个字段列表添加到Excel工作表
        except requests.exceptions.RequestException as e:
            print(f"请求失败: {e}")
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    output_file = f'kylinos_patch_{timestamp}.xlsx'
    workbook.save(output_file)
    workbook.close()

    # 将Excel文件转换为pandas DataFrame
    df = pd.read_excel(output_file, engine='openpyxl')
    df_introduction = df['详细介绍'].apply(clean_content)
    # 对'详细介绍'列进行分割
    for column_name, (start_pattern, end_pattern) in split_patterns.items():
        # 使用正则表达式分割字符串，并获取分隔符之间的文本
        df[column_name] = df_introduction.str.split(start_pattern, expand=True)[1].str.split(end_pattern, expand=True)[0]
    # 保存处理后的DataFrame回Excel文件
    df.to_excel(f'kylinos_patch_split_{timestamp}.xlsx', index=False)

if __name__ == "__main__":
    main()