from xlrd import open_workbook as xlrd_open_workbook
from requests import get as requests_get
import os

def saveImageByUrl(url, file_name):
    if url == '':
        return False
    try:
        response = requests_get(url)
        image = response.content
        with open(file_name, 'wb') as f:
            f.write(image)
        return True
    except Exception as e:
        print(f'ERROR: {e}')
    return False

if __name__ == "__main__":
    excel_path = 'C:\\Users\\哇咔咔\\Desktop\\2017级健康码\\17新测八闽健康码.xls'
    src = "腾讯收集表"
    data = xlrd_open_workbook(excel_path)
    table = data.sheet_by_index(0)
    nrow = table.nrows
    if src == "钉钉":
        header = table.row_values(0)
        idx_sno = header.index('工号')
        idx_sname = header.index('提交人')
        idx_location = header.index('所在地')
        idx_image = header.index('上传八闽健康码')
        for i in range(1, nrow):
            file_name = f"{table.cell(i, idx_sno).value}_{table.cell(i, idx_sname).value}_{table.cell(i, idx_location).value}"
            url = table.cell(i, idx_image).value
            extd_name = url[url.rindex('.')+1:]
            file_name = f"C:\\Users\\哇咔咔\\Desktop\\2017级健康码\\17新测八闽健康码\\{file_name}.{extd_name}"
            if saveImageByUrl(url, file_name):
                print(f'{i} / {nrow-1}  done')
            else:
                print(f'{i} / {nrow-1}  error!!!!!!!!!')
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
            file_name = f"C:\\Users\\哇咔咔\\Desktop\\2017级健康码\\17新测八闽健康码\\{file_name}.{extd_name}"
            if saveImageByUrl(url, file_name):
                print(f'{i} / {nrow-1}  done')
            else:
                print(f'{i} / {nrow-1}  error!!!!!!!!!')
