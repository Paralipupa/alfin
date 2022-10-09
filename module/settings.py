
import logging

logger = logging.getLogger('report')

ENCONING = 'utf-8'

FLD58_NAME = '0'
FLD58_BEG_DEBET = '2'
FLD58_TURN_DEBET = '4'
FLD58_TURN_CREDIT = '5'
FLD58_END_DEBET = '6'
FLD58_END_CREDIT = '7'

FLDPDN_NUMBER = '0'
FLDPDN_DATE = '1'
FLDPDN_NAME = '2'
FLDPDN_SUMMA = '4'
FLDPDN_PDN = '5'

FLDIRK_NUMBER = '0'
FLDIRK_PROC = '1'
FLDIRK_TARIF = '2'
FLDIRK_DATE = '3'
FLDIRK_NAME = '5'
FLDIRK_PASSPORT = '7'
FLDIRK_SUMMA = '13'
FLDIRK_PERIOD_COMMON = '14'
FLDIRK_PERIOD = '15'
FLDIRK_SUMMA_DEB_COMMOT = '16'
FLDIRK_SUMMA_DEB_MAIN = '17'
FLDIRK_SUMMA_DEB_PROC = '18'

PATT_NAME = '^\w+\s+\w+\s+\w+[вна|вич]\s*$'
PATT_FAMALY = '^\w+'
PATT_CURRENCY = '^-?\d{1,5}(?:[\.,]\d+)?$'
PATT_DOG_NAME ='^договор займа'
PATT_DOG_DATE='^[0-9]{1,2}\.[0-9]{2}\.20[1-9]{2}'
PATT_DOG_NUMBER='^20[0-9]{2}[0-9]{2}[0-9]{2}[0-9]{4}$'


OFLD_NAME = 'name'

                # "date": "18.08.2022 0:00:00",
                # "number": "202208180002",
                # "beg_debet": "9000",
                # "turn_debet": "9000",
                # "turn_credit": "",
                # "end_debet": "9000",
                # "found": true,
                # "pdn": "16.9",
                # "proc": "0.85",
                # "tarif": "Постоянный",
                # "passport": "2511 638379",
                # "period": "59",
                # "summa_deb_common": "10836",
                # "summa_deb_main": "9000",
                # "summa_deb_proc": "1836"
