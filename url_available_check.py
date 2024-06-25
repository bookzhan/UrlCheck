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
    # 定义含有可能的 URL 列名的列表
    url_column_names = ['URL', '网址', '链接', '网站']
    try:
        xls = pd.ExcelFile(file_path)

        directory = os.path.dirname(file_path)
        check_result_path = os.path.join(directory, 'check_result.xlsx')

        with pd.ExcelWriter(check_result_path) as writer:
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                found_url_column = None
                for col in url_column_names:
                    if col in df.columns:
                        found_url_column = col
                        break
                if found_url_column:
                    df['是否可访问'] = df[found_url_column].apply(check_website)
                df.to_excel(writer, sheet_name=sheet_name, index=False)

    except Exception as e:
        print(f"Error reading Excel file: {e}")

    print("url_available_check end")
