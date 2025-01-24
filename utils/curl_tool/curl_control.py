import re
import json
import pandas as pd


def _parse_curl(curl_data):

    url_match = re.search(r"curl '(.*?)'", curl_data)
    if url_match:
        url = str(url_match.group(1))
    else:
        return None

    header_matches = re.findall(r"-H '(.*?)'", curl_data)
    headers = {}
    for header in header_matches:
        key, value = header.split(": ", 1)
        headers[key] = value

    data_match = re.search(r"--data-raw '(.*)'", curl_data)
    data = json.loads(data_match.group(1)) if data_match else {}

    cookies = {}
    cookie_header = headers.get("cookie", "")
    cookie_matches = re.findall(r"([^=]+)=([^;]+);?", cookie_header)
    for key, value in cookie_matches:
        cookies[key] = value

    method_match = re.search(r"-X\s+([A-Z]+)\s+", curl_data)
    method = method_match.group(1) if method_match else "POST"

    return {
        "url": url,
        "headers": headers,
        "data": data,
        "cookies": cookies,
        "method": method
    }


def curl_to_excel(curl_data, excel_file):

    parsed_data = _parse_curl(curl_data)

    if parsed_data is not None:
        result = {
            "api-case-id": None,
            "项目模块": None,
            "用例标题": None,
            "依赖": {'code': 0, 'exp': [{'use': '', 'key': ''}], 'get': [''], 'set': ['']},
            "url": parsed_data["url"],
            "method": parsed_data["method"],
            "cookies": parsed_data["cookies"],
            "headers": parsed_data["headers"],
            "data_type": "json or form-data",
            "data": parsed_data["data"],
            "response": None,
            "预期结果": {'type': '==', 'key': '', 'value': ''},
            "备注": None
        }

        keys = list(result.keys())
        df = pd.DataFrame(columns=keys)
        df = df._append(result, ignore_index=True)
        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for i, col in enumerate(df.columns):
            column_len = max(df[col].astype(str).apply(len).max(), len(col)) + 2
            worksheet.set_column(i, i, column_len)
        writer.close()


if __name__ == '__main__':
    excel_file = '../../data/bilibili/bilibili.xlsx'
    data_curl = """
    curl 'https://api.bilibili.com/x/web-interface/wbi/search/all/v2?__refresh__=true&_extra=&context=&page=1&page_size=42&order=&duration=&from_source=&from_spmid=333.337&platform=pc&highlight=1&single_column=0&keyword=JKAI%E6%9D%B0%E5%87%AF&qv_id=Y9V1c9dxYPf7pyiU9kWoXFATpDNLn1Ag&ad_resource=5646&source_tag=3&web_location=1430654&w_rid=973f294a929c3226fce45a7d6865476c&wts=1724414473' \
  -H 'accept: application/json, text/plain, */*' \
  -H 'accept-language: zh-CN,zh;q=0.9' \
  -H 'cache-control: no-cache' \
  -H 'cookie: buvid3=B02B63CA-11EA-9E15-B9A1-37A9DED97B1539918infoc; b_nut=1716447039; _uuid=710F10AE71-F943-22B1-34BA-493E562DC1010542456infoc; buvid_fp=90d595f837b7269a7bcc284ac061c436; enable_web_push=DISABLE; home_feed_column=5; buvid4=A4E58401-177C-3EDE-6CFA-57B2D52D710645318-024052306-rZI1alEMf36x29EsxuGR3w%3D%3D; b_lsid=A3241E78_1917F1C72F0; header_theme_version=CLOSE; browser_resolution=2188-649' \
  -H 'origin: https://search.bilibili.com' \
  -H 'pragma: no-cache' \
  -H 'priority: u=1, i' \
  -H 'referer: https://search.bilibili.com/all?keyword=JKAI%E6%9D%B0%E5%87%AF&from_source=webtop_search&spm_id_from=333.1007&search_source=3' \
  -H 'sec-ch-ua: "Not)A;Brand";v="99", "Google Chrome";v="127", "Chromium";v="127"' \
  -H 'sec-ch-ua-mobile: ?0' \
  -H 'sec-ch-ua-platform: "macOS"' \
  -H 'sec-fetch-dest: empty' \
  -H 'sec-fetch-mode: cors' \
  -H 'sec-fetch-site: same-site' \
  -H 'user-agent: Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36'
   """
    curl_to_excel(curl_data=data_curl, excel_file=excel_file)
    print(f"Excel 文件已生成: {excel_file}")
"""
charles与浏览器copy导出的curl格式不一样，需要兼容正则

url = re.search(r'curl\s.*?["\'](.*?)["\']', curl_command).group(1)
        method = re.search(r'-X (\w+)', curl_command)
        method = method.group(1) if method else 'GET'
        headers = re.findall(r'-H ["\'](.*?)["\']', curl_command)
        data = re.search(r'--data-binary ["\'](.*?)["\']', curl_command)
        if not data:
            data = re.search(r'--data-raw ["\'](.*?)["\']', curl_command)
        data = data.group(1) if data else ''

"""