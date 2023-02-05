
import logging

logger = logging.getLogger('report')

ENCONING = 'utf-8'

PATT_NAME = '(?:на|ич)\s*$'
PATT_FAMALY = '^\w+'
PATT_CURRENCY = '^-?\d{1,8}(?:[\.,]\d+)?$'
PATT_PROC = '^\d{1,3}(?:[\.,]\d+)?$'
PATT_PDN = '^\d{1,3}(?:[\.,]\d+)?$'
PATT_TARIF = '(?:постоянный|старт|31|24)$'
PATT_PERIOD = '^\d{2,4}$'
PATT_COUNT_DAYS = '^\d{2,4}$'
PATT_DOG_TYPE = '^ЯЯ'
PATT_DOG_NAME ='^договор займа'
PATT_DOG_DATE='^[0-9]{1,2}\.[0-9]{2}\.20[0-9]{2}'
PATT_DOG_NUMBER='^(?:ON)?20[0-9]{2}[0-9]{2}[0-9]{2}[0-9]{4}$|^(?:ON)?[a-zA-Zа-яА-Я0-9]{1,2}[0-9]{6}[0-9]{4}\s*$'
PATT_DOG_PLAT='Обороты за '

SQL_CONNECT = {
    'dsn' : 'sqlserverdatasource',
    'port' :'1433',
    'database' : 'ArchiCreditW',
    'user':'sa', 
    'password': 'Raideff86reps$1'}




