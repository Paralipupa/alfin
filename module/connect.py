import pyodbc
import sys, time, os
import pandas as pd

sys.path.append("alfin")
from module.settings import *

logger = logging.getLogger(__name__)


class SQLServer:
    def __init__(self):
        self.connection = None
        self.set_connection()

    def set_connection(self):
        con_string = "DSN=%s;PORT=%s;UID=%s;PWD=%s;DATABASE=%s;" % (
            SQL_CONNECT["dsn"],
            SQL_CONNECT["port"],
            SQL_CONNECT["user"],
            SQL_CONNECT["password"],
            SQL_CONNECT["database"],
        )
        # con_string = "DRIVER={{FreeTDS}};SERVER={server};DATABASE={database};UID={user};PWD={password}".format(
        #     server=SQL_CONNECT["server"],
        #     database=SQL_CONNECT["database"],
        #     user=SQL_CONNECT["user"],
        #     password=SQL_CONNECT["password"],
        # )

        try:
            self.connection = pyodbc.connect(con_string)
            return True
        except Exception as ex:
            logger.error(f"{con_string}\n {ex}")
        return False

    def get_orders(self, numbers: list = ["0"]):
        mSQL = "SELECT o.ID, o.MAINPERCENT, o.DAYSQUANT, o.NUMBER, cast(c.[CREATIONDATETIME] as DateTime) as CREATEDATE,"
        mSQL += "c.FULLNAME, o.POID, p.NAME, o.LOANCOSTALL,c.DOCS,c.DOCNUM,cast(c.[DOCBEGINDATE] as DateTime) as DOCDATE"
        # mSQL += "c.DOCCONTENT "
        mSQL += " FROM [Orders] o "
        mSQL += " INNER JOIN [CLIENTS] c ON c.[ID]=o.[CLIENTID]"
        mSQL += " INNER JOIN [PERCENT_OPTIONS] p ON o.[POID]=p.[ID]"
        mSQL += " WHERE o.[NUMBER] in ('{}')".format("','".join(numbers))
        try:
            cursor = self.connection.cursor()
            cursor.execute(mSQL)
            results = [list(x) for x in cursor.fetchall()]
            keys = [f"{x[3]}" for x in results]
            data = dict(zip(keys, results))
            return data
        except Exception as ex:
            logger.error(f"{ex}")
        return []

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
        mSQL = "SELECT o.[ID], o.[NUMBER] AS 'NUMBER_ORDER', c.[COSTALL], cast(c.[CREATIONDATETIME] as DateTime) as CREATEDATE,"
        mSQL += "c.[ENABLED], c.[KIND],c.[NUMBER] AS 'NUMBER_PAYMENT'  "
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

    def get_data_clients(self, letter):
        mSQL = f"""SELECT 
	   c.[ID]
      ,[FULLNAME]
      ,[BIRTHDATE]
      ,[DOCS]
      ,[DOCNUM]
      ,[DOCBEGINDATE]
      ,[DOCCONTENT]
      ,[ADDRESS_REG]
      ,[ADDRESS_FACT]
      ,[OLDFIO]
      ,[PHONE]
      ,[MOBILEPHONE]
      ,[BLACKLIST]
      ,[BIRTHPLACE]
      ,[CREDITHISTORY]
  FROM [ArchiCredit].[dbo].[CLIENTS] c
  LEFT JOIN dbo.CLIENTS_CREDIT_HISTORY h
  ON c.id=h.CLIENTID
  WHERE (Len(FULLNAME) > 8)  AND c.[NAME] Like '{letter}%' ORDER BY NAME,NAME1,NAME2        """
        cursor = self.connection.cursor()
        cursor.execute(mSQL)
        results = [list(x) for x in cursor.fetchall()]
        return results


if __name__ == "__main__":
    ascii_uppercase = 'АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ'
    q = SQLServer()
    start= time.time()
    if q.set_connection():
        for letter in ascii_uppercase:
            file_name = f"clients/{letter}"
            if not os.path.exists(f"{file_name}.csv"):
                print(letter, flush=True, end="")
                data = q.get_data_clients(letter)
                idents = set([x[0] for x in data])
                items = list()
                for item in data:
                    if item[0] in idents:
                        items.append(item)
                        idents.remove(item[0])
                client_df = pd.DataFrame([x[:-1] for x in items])
                history_df = pd.DataFrame([(x[0], x[-1]) for x in items])
                client_df.to_csv(
                    f"{file_name}.csv", index=False, header=False, sep=";"
                )
                history_df.to_csv(
                    f"{file_name}_history.csv", index=False, header=False, sep=";"
                )
        q.connection.close()
    print(f"\nOK {(time.time()-start):10.2f} сек")
