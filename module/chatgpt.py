import os, sys
import time
import pandas as pd
import pyodbc
from concurrent.futures import ProcessPoolExecutor, ThreadPoolExecutor, as_completed

sys.path.append("alfin")
from module.settings import *

logger = logging.getLogger(__name__)


class SQLServer:
    def __init__(self):
        self.connection = None

    def set_connection(self):
        con_string = "DSN=%s;PORT=%s;UID=%s;PWD=%s;DATABASE=%s;" % (
            SQL_CONNECT["dsn"],
            SQL_CONNECT["port"],
            SQL_CONNECT["user"],
            SQL_CONNECT["password"],
            SQL_CONNECT["database"],
        )
        try:
            self.connection = pyodbc.connect(con_string)
            return True
        except Exception as ex:
            logger.error(f"{con_string}\n {ex}")
        return False

    def get_data_clients(self, letter, debug=False):
        mSQL = f"""SELECT 
        {'TOP 10' if debug else ''}
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

    def close_connection(self):
        if self.connection:
            self.connection.close()


def download_clients():

    def process_letter(letter):
        q = SQLServer()
        if q.set_connection():
            file_name = f"clients/{letter}"
            if not os.path.exists(f"{file_name}.csv"):
                data = q.get_data_clients(letter)
                idents = set([x[0] for x in data])
                items = list()
                for item in data:
                    if item[0] in idents:
                        items.append(item)
                        idents.remove(item[0])
                client_df = pd.DataFrame([x[:-1] for x in items])
                history_xml = [[x[0], x[1], x[-1]] for x in items]
                client_df.to_csv(f"{file_name}.csv", index=False, header=False, sep=";")
                path = f"clients/history/{letter}"
                os.makedirs(path, exist_ok=True)
                for his in history_xml:
                    if his[-1] is not None:
                        name = "{0}_{1}".format(his[0], his[1])
                        with open(f"{path}/{name}.xml", mode="w", encoding="windows-1251") as f:
                            f.write(his[-1])
                print(letter)
            q.close_connection()

    print("Начало ")
    ascii_uppercase = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЫЭЮЯ"
    start = time.time()
    os.makedirs("clients", exist_ok=True)
    os.makedirs("clients/history", exist_ok=True)

    # with ProcessPoolExecutor(max_workers=4) as executor:
    #     executor.map(process_letter, ascii_uppercase)
    with ThreadPoolExecutor() as executor:
        futures = [
            executor.submit(process_letter, letter) for letter in ascii_uppercase
        ]
        for future in as_completed(futures):
            future.result()
    # _ = list(map(process_letter, ascii_uppercase))

    print(f"\nОк ({time.strftime('%H:%M:%S', time.gmtime(time.time()-start))} сек)")


if __name__ == "__main__":
    download_clients()
