
import logging

logger = logging.getLogger('report')

ENCONING = 'utf-8'

PATT_NAME = '(?:на|ич)\s*$'
PATT_FAMALY = '^\w+'
PATT_CURRENCY = '^-?\d{1,8}(?:[\.,]\d+)?$'
PATT_PROC = '^\d{1,3}(?:[\.,]\d{1,3})$'
PATT_PDN = '^\d{1,3}(?:[\.,]\d+)$'
PATT_TARIF = '(?:постоянный|старт)$'
PATT_PERIOD = '^\d{2,4}$'
PATT_COUNT_DAYS = '^\d{2,4}$'
PATT_DOG_NAME ='^договор займа'
PATT_DOG_DATE='^[0-9]{1,2}\.[0-9]{2}\.20[1-9]{2}'
PATT_DOG_NUMBER='^20[0-9]{2}[0-9]{2}[0-9]{2}[0-9]{4}$'



