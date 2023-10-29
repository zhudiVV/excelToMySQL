import openpyxl
import pymysql

class ExcelUtils(object):
    def __init__(selfself):
        return

    def get_conn(self):
        try:
            conn = pymysql.connect(host="localhost",
                                   port=3306,
                                   user="root",
                                   password="Star2035!",
                                   charset="utf8mb4")
        except:
            pass
        return conn
    def export_xls(self):
        conn = self.get_conn()

        # 创建游标对象
        cursor = conn.cursor()
        # 选择数据库
        conn.select_db("nbjy_admin_db")


        # 打开文件
        workbook = openpyxl.load_workbook('/Users/zhudi/Downloads/user.xlsx')

        # workbook_json = json.dumps(workbook)
        # print(repr(workbook))
        # print(workbook_json)
        ws = workbook['user']

        header = []
        for index, row in enumerate(ws.rows):
            if index == 0:
                for cell in row:
                    header.append(cell.value)
            else:
                result1 = ','.join('%s' % item for item in header)


                varList = []
                for i,cell in enumerate(row):
                    # str2 = \' + cell.value + '\''
                    # if isinstance(cell.value, uni)
                    if cell.value:
                        varList.append(str(cell.value).replace(',', '-').replace('，', '-'))
                    else:
                        varList.append('')


                result2 = ','.join('\'%s\'' % item for item in varList)
                # print(result1)
                # print(result2)
                sql_format = "INSERT INTO user2 ({}) VALUES ({});"
                # sql = "INSERT INTO user2 (" + result1 + ") VALUES( " + result2 +")"
                sql = sql_format.format(result1, result2)
                print(sql)



                cursor.execute(sql)
                conn.commit()

        # 关闭游标和连接
        cursor.close()
        conn.close()





if __name__ == '__main__':
    client = ExcelUtils()
    client.export_xls()