import pyodbc
con_string = "DSN=%s;PORT=%s;UID=%s;PWD=%s;DATABASE=%s;" % (
    SQL_CONNECT["dsn"],
    SQL_CONNECT["port"],
    SQL_CONNECT["user"],
    SQL_CONNECT["password"],
    SQL_CONNECT["database"],
)

try:
    self.connection = pyodbc.connect(con_string)
except Exception as ex:
    pass
