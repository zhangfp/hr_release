#coding=utf-8
import os,json
from pymongo import MongoClient
from openpyxl import load_workbook
from openpyxl import Workbook
import xlrd



#连接数据库
client=MongoClient('localhost',27017)
mongodb=client.test
hr=mongodb.hr
hr_keys=mongodb.hr_keys
project_map_table=mongodb.project_map


def hr_template(hr_template_file):
    if os.path.exists(hr_template_file):
        data=xlrd.open_workbook(hr_template_file)
        table=data.sheets()[0]
        #读取excel第一行数据作为存入mongomongodb的字段名
        rowstag=table.row_values(0)
        #连接数据库
        hr.drop()
        hr_keys.drop()

        wb = load_workbook(filename = hr_template_file,read_only=True)
        sheets = wb.get_sheet_names()   
        ws = wb[sheets[0]]
        rows = ws.rows

        content = []

        rowstag_dict = {
            'excel_keys': rowstag
        }
        hr_keys.insert(rowstag_dict)


        for row in rows:
            line = [col.value for col in row]
            returnData=json.dumps(dict(zip(rowstag,line)))
            #print returnData
            #通过编解码还原数据
            returnData=json.loads(returnData)
            content.append(returnData)
            #hr.insert(returnData)
            
        hr.insert_many(content)

def main():
    currFolder = os.path.dirname(os.path.realpath(__file__));
    UPLOAD_FOLDER = os.path.join(currFolder, 'static/uploads')
    filename = "work_time_template.xlsx"
    upload_file_path = UPLOAD_FOLDER + '/' + filename

    hr_template(upload_file_path)

if __name__ == '__main__':
    main()
