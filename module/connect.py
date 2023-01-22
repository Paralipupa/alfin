import pyodbc
from module.settings import *

class SQLServer():
    
    def __init__(self):
        self.connection = None
        self.set_connection()

    def set_connection(self):
        con_string = 'DSN=%s;PORT=%s;UID=%s;PWD=%s;DATABASE=%s;' % (
            SQL_CONNECT['dsn'], SQL_CONNECT['port'], SQL_CONNECT['user'],
            SQL_CONNECT['password'], SQL_CONNECT['database'])
        try:
            self.connection = pyodbc.connect(con_string)
        except Exception as ex:
            pass

    def get_orders(self, numbers: list = ['0']):
        mSQL = "SELECT o.ID, o.MAINPERCENT, o.DAYSQUANT, o.NUMBER, o.CREATIONDATETIME, c.FULLNAME, o.POID, p.NAME"
        mSQL = mSQL + " FROM [Orders] o "
        mSQL = mSQL + " INNER JOIN [CLIENTS] c ON c.[ID]=o.[CLIENTID]"
        mSQL = mSQL + " INNER JOIN [PERCENT_OPTIONS] p ON o.[POID]=p.[ID]"
        mSQL = mSQL + " WHERE o.[NUMBER] in ('{}')".format("','".join(numbers))
        # mSQL = mSQL + " ORDER BY o.ID DESC"
        cursor = self.connection.cursor()
        cursor.execute(mSQL)
        results = cursor.fetchall()
        keys = [f"{x[3]}" for x in results]
        data = dict(zip(keys, results))
        return data

    def get_orders_frost(self, orders: str='-1'):
        mSQL = "SELECT TOP 10 o.[ORDERID], o.[CREATIONDATETIME] "
        mSQL = mSQL + " FROM [Order_Frost] o "
        # mSQL = mSQL + " WHERE o.[ORDERID] in ({})".format(orders)
        mSQL = mSQL + " ORDER BY o.ID DESC"
        cursor = q.connection.cursor()
        cursor.execute(mSQL)
        results = cursor.fetchall()
        keys = [f"{x[0]}" for x in results]
        data = dict(zip(keys, results))
        return data

    def get_orders_payments(self, orders: str='-1'):
        mSQL = "SELECT TOP 10 o.[ORDERID], o.[CREATIONDATETIME],o.[ENABLED], o.[KIND], o.[COSTALL] "
        mSQL = mSQL + " FROM [Order_Payment] o "
        # mSQL = mSQL + " WHERE o.[ORDERID] in ({})".format(orders)
        mSQL = mSQL + " ORDER BY o.ID DESC"
        cursor = q.connection.cursor()
        cursor.execute(mSQL)
        results = cursor.fetchall()
        keys = [f"{x[0]}" for x in results]
        data = dict(zip(keys, results))
        return data
    
    def get_data_from_archi(self, numbers: list):
        data = self.get_orders(numbers)
        # q.get_orders_frost().values()
        # q.get_orders_payments().values()
        self.connection.close()
        return data

if __name__ == '__main__':
    q = SQLServer()
    if q.set_connection():
        [print(x) for x in q.get_orders().values()]
        [print(x) for x in q.get_orders_frost().values()]
        [print(x) for x in q.get_orders_payments().values()]
        q.connection.close()
