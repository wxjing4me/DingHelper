from xlrd import open_workbook as xlrd_open_workbook
from requests import get as requests_get
import os

SOURCE_EXCEL_PATH = "C:\\Users\\哇咔咔\\Desktop\\每日健康打卡\\2016级健康码.xlsx"
TARGET_FOLDER = "C:\\Users\\哇咔咔\\Desktop\\每日健康打卡\\2016级健康码\\"
SOURCE_TYPE = "钉钉"  # 腾讯收集表 or 钉钉

def saveImageByUrl(url, file_name):
    if url == '':
        return False
    try:
        response = requests_get(url, timeout=2)
        image = response.content
        with open(file_name, 'wb') as f:
            f.write(image)
        return True
    except Exception as e:
        print(f'ERROR: {e}')
    return False

if __name__ == "__main__":
    data = xlrd_open_workbook(SOURCE_EXCEL_PATH)
    table = data.sheet_by_index(0)
    nrow = table.nrows
    if SOURCE_TYPE == "钉钉":
        header = table.row_values(0)
        idx_sno = header.index('工号')
        idx_sname = header.index('提交人')
        idx_location = header.index('当前时间,当前地点')
        idx_image = header.index('请上传你的健康识别码')
        for i in range(1, nrow):
            try:
                string = table.cell(i, idx_location).value
                location = string[[i for i,x in enumerate(string) if x == '"' ][2]+1:][:6]
                file_name = f"{table.cell(i, idx_sno).value}_{table.cell(i, idx_sname).value}_{location}"
                url = table.cell(i, idx_image).value
                extd_name = url[url.rindex('.')+1:]
                file_name = f"{TARGET_FOLDER}{file_name}.{extd_name}"
                saveImageByUrl(url, file_name)
                print(f'{i} / {nrow-1}  done')
            except Exception as e:
                print(f'{i} / {nrow-1} {table.cell(i, idx_sname).value}  error!!!!!!!!! - {e}')
    else:
        # src == "腾讯收集表"
        header = table.row_values(0)
        idx_sno = header.index('学号')
        idx_sname = header.index('姓名')
        idx_image = header.index('八闽健康码')
        for i in range(1, nrow):
            file_name = f"{table.cell(i, idx_sno).value}_{table.cell(i, idx_sname).value}"
            link = table.hyperlink_map.get((i, idx_image))
            url = '' if link is None else link.url_or_path
            extd_name = url[url.rindex('=')+1:]
            file_name = f"{TARGET_FOLDER}{file_name}.{extd_name}"
            if saveImageByUrl(url, file_name):
                print(f'{i} / {nrow-1}  done')
            else:
                print(f'{i} / {nrow-1}  error!!!!!!!!!')
