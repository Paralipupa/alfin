import pyodbc
from module.settings import *


class SQLServer:
    def __init__(self):
        self.connection = None
        self.set_connection()

    # install postgresql
    # https://www.postgresql.org/download/linux/ubuntu/
    #
    # /etc/odbc.ini:
    # [sqlserverdatasource]
    # Driver = FreeTDS
    # Description = ODBC connection via FreeTDS
    # Trace = No
    # Servername = sqlserver
    # Database = ArchiCreditW
    # TDS_Version = 7.4
    def set_connection(self):
        con_string = "DSN=%s;PORT=%s;UID=%s;PWD=%s;DATABASE=%s;" % (
            SQL_CONNECT["dsn"],
            SQL_CONNECT["port"],
            SQL_CONNECT["user"],
            SQL_CONNECT["password"],
            SQL_CONNECT["database"],
        )
        # con_string = "DRIVER={{ODBC Driver 18 for SQL Server}};SERVER={server};DATABASE={database};UID={user};PWD={password}".format(
        #     server=SQL_CONNECT["server"],
        #     database=SQL_CONNECT["database"],
        #     user=SQL_CONNECT["user"],
        #     password=SQL_CONNECT["password"],
        # )

        try:
            self.connection = pyodbc.connect(con_string)
        except Exception as ex:
            pass

    def get_orders(self, numbers: list = ["0"]):
        mSQL = "SELECT o.ID, o.MAINPERCENT, o.DAYSQUANT, o.NUMBER, cast(c.[CREATIONDATETIME] as DateTime) as CREATEDATE, c.FULLNAME, o.POID, p.NAME, o.LOANCOSTALL"
        mSQL = mSQL + " FROM [Orders] o "
        mSQL = mSQL + " INNER JOIN [CLIENTS] c ON c.[ID]=o.[CLIENTID]"
        mSQL = mSQL + " INNER JOIN [PERCENT_OPTIONS] p ON o.[POID]=p.[ID]"
        mSQL = mSQL + " WHERE o.[NUMBER] in ('{}')".format("','".join(numbers))
        cursor = self.connection.cursor()
        cursor.execute(mSQL)
        results = [list(x) for x in cursor.fetchall()]
        keys = [f"{x[3]}" for x in results]
        data = dict(zip(keys, results))
        return data

    def get_orders_frost(self, numbers: str = "-1"):
        mSQL = "SELECT o.[ID], o.NUMBER, cast(c.[CREATIONDATETIME] as DateTime) as CREATEDATE "
        mSQL = mSQL + " FROM [Orders] o "
        mSQL = mSQL + " INNER JOIN [Order_Frost] c ON c.[ORDERID]=o.[ID]"
        mSQL = mSQL + " WHERE o.[NUMBER] in ('{}')".format("','".join(numbers))
        mSQL = mSQL + " ORDER BY o.ID DESC"
        cursor = self.connection.cursor()
        cursor.execute(mSQL)
        results = [list(x) for x in cursor.fetchall()]
        keys = [f"{x[1]}" for x in results]
        data = dict(zip(keys, results))
        return data

    def get_orders_payments(self, numbers: str = "-1"):
        mSQL = "SELECT o.[ID], o.[NUMBER], c.[COSTALL], cast(c.[CREATIONDATETIME] as DateTime) as CREATEDATE,c.[ENABLED], c.[KIND] "
        mSQL = mSQL + " FROM [Orders] o "
        mSQL = mSQL + " INNER JOIN [Order_Payment] c ON c.[ORDERID]=o.[ID]"
        mSQL = mSQL + " WHERE o.[NUMBER] in ('{}')".format("','".join(numbers))
        mSQL = mSQL + " ORDER BY o.ID DESC"
        cursor = self.connection.cursor()
        cursor.execute(mSQL)
        results = [list(x) for x in cursor.fetchall()]
        data = {}
        for item in results:
            data.setdefault(item[1], [])
            data[item[1]].append(item)
        return data

    def get_data_from_archi(self, numbers: list):
        data = {}
        data["order"] = self.get_orders(numbers)
        data["frost"] = self.get_orders_frost(numbers)
        data["payment"] = self.get_orders_payments(numbers)
        self.connection.close()
        return data


if __name__ == "__main__":
    q = SQLServer()
    if q.set_connection():
        [print(x) for x in q.get_orders().values()]
        [print(x) for x in q.get_orders_frost().values()]
        [print(x) for x in q.get_orders_payments().values()]
        q.connection.close()
