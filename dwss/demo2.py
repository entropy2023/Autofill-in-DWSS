# -*- coding: utf-8 -*-
from lxml import etree
import requests
import pandas as pd

# 指定你的 Excel 文件路径

excel_file_path = r'C:\Users\zhang\Desktop\test\dwss\3_test.xlsx'

# 使用 pandas 读取 Excel 文件，指定 openpyxl 作为引擎
df = pd.read_excel(excel_file_path, engine='openpyxl')

# 遍历 DataFrame 中的每一行
for index, row in df.iterrows():
    if index == 0:
        continue
    #将列表中的时间参数转化为12-07-2024 10:20的格式
    INSPECTION_DATE=row.to_dict().get('The Following works will be ready for inspection on').strftime('%d-%m-%Y') + ' ' + row.to_dict().get('Unnamed: 4').strftime('%H:%M')
    print(INSPECTION_DATE)
    LOCATION=row.to_dict().get('Unnamed: 5')

    DESC_OF_WORKS=row.to_dict().get('Unnamed: 6')

    WORKS_PROPOSED=row.to_dict().get('Unnamed: 7')

    CONTACT_PERSON=row.to_dict().get('Unnamed: 8')          #表格所有项均不能为空

    RFI_ITAP_QSSP=row.to_dict().get('Unnamed: 13')

    MAIN_WORK_CATEGORIES=row.to_dict().get('Unnamed: 11')

    SUB_WORK_CATEGORIES=row.to_dict().get('Unnamed: 12')     #从excel表中获取数据，注意excel中的ITAP Code,Sub-works Category字段一定要和网页元素保持一致，不能用简写(比如示例中给出的SF2)，大小写也一定要正确(比如网页中是Others,而excel表中的示例是other)。这段自动化代码本质上是对保存数据的网站发送网络请求，这些字段也一定要与网页元素保持一致，必须一模一样！

    # 指定你的 HTML 文件路径
    html_file_path = './id.html'

    # 读取并解析本地 HTML 文件
    with open(html_file_path, 'r', encoding='utf-8') as file:
        html_content = file.read()
        tree = etree.HTML(html_content)

    # 使用 XPath 表达式找到特定的 ul 标签
    ul_xpath = '//*[@id="select2-RFI_ITAP_QSSP_ID-results"]'
    ul_element = tree.xpath(ul_xpath)[0]  # 假设只有一个这样的 ul

    # 遍历 ul 下的所有 li 标签
    li_elements = ul_element.xpath('.//li')

    for li in li_elements:          #获取MAIN_WORK_CATEGORIES_NO
        if li == li_elements[0]:
            continue

        if RFI_ITAP_QSSP == li.xpath('./text()')[0]:
            RFI_ITAP_QSSP_ID = li.get('id').split('-')[4]    #

    if MAIN_WORK_CATEGORIES == 'Architectural':   #获取MAIN_WORK_CATEGORIES_NO和SUB_WORK_CATEGORIES_ID
        MAIN_WORK_CATEGORIES_NO = '26'
        SUB_WORK_CATEGORIES_ID = 'N/A'
    else:
        MAIN_WORK_CATEGORIES_NO = '27'

        if SUB_WORK_CATEGORIES == 'Excavation and earthwork':
            SUB_WORK_CATEGORIES_ID = '31'
        elif SUB_WORK_CATEGORIES == 'Others':
            SUB_WORK_CATEGORIES_ID = '35'
        elif SUB_WORK_CATEGORIES == 'Piling /Foundation':
            SUB_WORK_CATEGORIES_ID = '32'
        elif SUB_WORK_CATEGORIES == 'Reinforced concrete':
            SUB_WORK_CATEGORIES_ID = '33'
        else:
            SUB_WORK_CATEGORIES_ID = '34'

    cookies = {
        'eirs': '!ahhg2Qjgd/VCmSwAz5JTUhgTkMNW57ytrhSVXwGaupbk/6tComrHpFyr4VBBFUFoXdd68rUNlPTIJtg=',
        'ASP.NET_SessionId': 'yxg4ltfy2ilt1z2pvu2ovmzg',
    }    #你验证后的cookies,记得修改

    headers = {
        'Accept': '*/*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        # 'Cookie': 'eirs=!txBAluFQT/ZEeNEAz5JTUhgTkMNW5wcm1zoENWvrAQaNK4Of+uEMR45VxIG+EUZBtVwypNpqWRGCfUs=; ASP.NET_SessionId=1knlnznnsupolit5tx5y3chi',
        'Origin': 'https://dwss.archsd.gov.hk',
        'Pragma': 'no-cache',
        'Referer': 'https://dwss.archsd.gov.hk/DWSS/RFI/CreateRFI',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36 Edg/127.0.0.0',
        'X-Requested-With': 'XMLHttpRequest',
        'sec-ch-ua': '"Not)A;Brand";v="99", "Microsoft Edge";v="127", "Chromium";v="127"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }     #请求头

    data = 'MAIN_CONTRACTOR_COMPANY_ID=1063&ROLE_NAME=MC&PRJ_PROJECT_ID=1205005&DWG_SKETCH_NO=&SERIAL_NO=&MAIN_WORK_CATEGORIES=' + MAIN_WORK_CATEGORIES_NO + '&SUB_WORK_CATEGORIES_ID=' + SUB_WORK_CATEGORIES_ID + '&RFI_ITAP_QSSP_ID=' + RFI_ITAP_QSSP_ID + '&INSPECTION_DATE=' + INSPECTION_DATE + '&LOCATION=' + LOCATION + '&LOCATION_ID=-1&DESC_OF_WORKS=' + DESC_OF_WORKS + '&WORKS_PROPOSED=' + WORKS_PROPOSED + '&DWG_SKETCH_FILE_NAME=&RFIAdditionalDocumentPDF_FILE_NAME=&To=CE%2FCOW&DELIVERY_MODE=D&CONTACT_PERSON=' + CONTACT_PERSON + '&CONTACT_NUMBER=69339311&PROJECT_SPECIFIC_CODE=SSM518&Old_RFI_ID=&BIM='

    response = requests.post('https://dwss.archsd.gov.hk/DWSS/RFI/Save', cookies=cookies, headers=headers, data=data)

    print(response.text)



