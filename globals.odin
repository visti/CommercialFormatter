package komm_fmt

import "core:time"


NOW := time.now()
CONVERT :: "/bin/python3 /mnt/d/komm_fmt/lib/convert.py *xls*"
DELETE_COLS :: "/bin/python3 /mnt/d/komm_fmt/lib/delete_columns.py "
REJECTDIR :: "/mnt/c/Users/eva/Gramex/Rapporteringer - Documents/Afviste_linjer_kom_land"
SEPARATOR := ";"
REJECTFILE := "rejected.csv"
EMPTY: string = ""
// ERROR MESSAGES
NOTVALIDSTATION :: "ERROR: No valid station selected as argument."
