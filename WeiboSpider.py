"""
Created on Sun Oct 18 14:08:00 2020

@author: Weixiang
"""

import requests,re 
import time,datetime
import xlrd,xlwt,json
from xlutils.copy import copy 

def get_title(path,month,day):
    # 设置参数 
    setdate = datetime.datetime.strptime(str(month)+"-"+str(day),'%m-%d').date()
    since_id = ''
    count = 0
    excelCount = 2
    flag = 10

    # 开始运行
    while True:   
        count += 1   
        headers = { 
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.100 Safari/537.36' 
        }        
        if since_id == '':
            api_url = 'https://m.weibo.cn/api/container/getIndex?type=uid&value=1698264705&containerid=1076031698264705'
        else:
            api_url = 'https://m.weibo.cn/api/container/getIndex?type=uid&value=1698264705&containerid=1076031698264705&since_id=' + str(since_id) 
        rep = requests.get(url=api_url, headers=headers).json()['data'] # 获取 ID 值并写入列表 comment_ID 中 
        # 获取since_id
        later_since_id = since_id
        since_id = rep['cardlistInfo']['since_id']
        
        Textlist = []
        index = -1
        while True:
            index += 1
            if index >= len(rep['cards']):
                break
            if 'mblog' not in rep['cards'][index].keys():
                continue
            # 时间
            created_at = rep['cards'][index]['mblog']['created_at']
            # 点赞数
            attitudes_count = rep['cards'][index]['mblog']['attitudes_count']
            # 地址
            url = 'https://m.weibo.cn/detail/' + rep['cards'][index]['mblog']['id']
            # 内容
            text = rep['cards'][index]['mblog']['text']
            if re.findall('>全文</a>',text,re.S) != []:
                html = requests.get(url,headers=headers).text
                jsondata = re.findall('render_data = \[(.*?)\]\[0\]',html,re.S)[0]
                pagejson = json.loads(jsondata)
                text = pagejson['status']['text']
            # [日期，点赞数，链接地址，内容]
            Textlist.append([created_at,attitudes_count,url,text])
            # print([created_at,attitudes_count,url,text])
        # print(str(later_since_id) + "-->" + str(since_id))

        # 写入excel
        wb = xlrd.open_workbook(path)
        # 将操作文件对象拷贝，变成可写的workbook对象
        workbook = copy(wb)
        # 获得第一个sheet的对象
        worksheet = workbook.get_sheet('sheet1')

        for i in Textlist:
            for n in range(0,len(i)):
                worksheet.write(excelCount,n, label = i[n])
            excelCount += 1
        # 保存
        workbook.save(path)
        print("Saved successfully！")
        
        # 循环终止条件
        if count > flag:
            date_p = datetime.datetime.strptime(created_at,'%m-%d').date()
            if date_p < setdate:
                break
    
    print("Finish!")
        
if __name__ == '__main__':    
    path = "weiboSpider.xls"
    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding = 'utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet("sheet1")
    worksheet.write(0,0, label = '发布时间')
    worksheet.write(0,1, label = '点赞量')
    worksheet.write(0,2, label = '链接地址')
    worksheet.write(0,3, label = '内容')
    workbook.save(path)
    
    # 参数说明：[excel文件地址，月，日]
    # 如 get_title(path,10,1) 表示
    # 爬取10月1日之后的微博，将其保存在path路径下的excel文件中
    # excel文件的后缀必须是 .xls
    get_title(path,10,1)

    
    