import xlrd
import json

# 第1行默认为表头
START_ROW = 1 # default:1

# 第1,2列默认为学号、姓名
START_COL = 2 # default:2

def testExcel(excel_path):
    excel = xlrd.open_workbook(excel_path)
    try:    
        excel = xlrd.open_workbook(excel_path)
        return True
    except:
        return False

def readExcel(excel_path):
    stuDatas = []
    try:
        data = xlrd.open_workbook(excel_path)
        table = data.sheet_by_index(0)
    except:
        return stuDatas
    nrow = table.nrows
    header = table.row_values(0)
    # print(header)
    for i in range(START_ROW, nrow):
        stu = {}
        stuInfo = [table.cell(i,0).value, table.cell(i,1).value]
        stu['info'] = stuInfo
        stuData = {}
        for j in range(START_COL, len(header)):
            value = table.cell(i,j).value
            try:
                value = json.loads(value)
                value = value[1:4]
            except Exception as e:
                print(e)
                pass
            stuData[header[j]] = value
        stu['data'] = stuData
        stuDatas.append(stu)
    return stuDatas