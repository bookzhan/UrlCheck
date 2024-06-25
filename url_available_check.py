import os

import requests
import pandas as pd


def check_website(_url):
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
        }
        response = requests.get(_url, headers=headers)
        if response.status_code == 200 and len(response.content) > 0:
            print("可访问:" + str(_url))
            return '可访问'
        else:
            print("不可访问:" + str(_url))
            return '不可访问'
    except requests.exceptions.RequestException as e:
        print("发生异常:" + str(_url) + "\n" + str(e))
        return '发生异常'


if __name__ == '__main__':
    print("url_available_check start")
    # Excel的路径,可以是相对路径,也可以是绝对路径
    file_path = "网站.xlsx"
    # 表名,不输入默认第一张表
    sheet_name = ""

    try:
        # 使用pandas读取Excel文件
        if len(sheet_name) > 0:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file_path)

        if 'URL' in df.columns:
            url_list = df['URL'].tolist()

            result_urls = []
            for url in url_list:
                result_type = check_website(url)
                result_urls.append({"URL": url, "Result": result_type})

            result_df = pd.DataFrame(result_urls)
            directory = os.path.dirname(file_path)

            check_result_path = os.path.join(directory, 'check_result.xlsx')
            if os.path.exists(check_result_path):
                os.remove(check_result_path)
            if not result_df.empty:
                result_df.to_excel(check_result_path, index=False)

        else:
            print("The specified Excel file does not contain a 'URL' column.")

    except Exception as e:
        print(f"Error reading Excel file: {e}")

    print("url_available_check end")
