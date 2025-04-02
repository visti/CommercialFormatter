package komm_fmt

import "core:time"

REPLACES :: []string{" - ABC PowerHit"}
NOW := time.now()
CONVERT :: "/bin/python3 /mnt/d/commercial_formatter/lib/convert.py *xls*"
MEMORY_LOG :: "memory_leak.log"
DELETE_COLS :: "/bin/python3 /mnt/d/commercial_formatter/lib/delete_columns.py "
REJECTDIR :: "/mnt/c/Users/eva/Gramex/Rapporteringer - Documents/Afviste_linjer_kom_land"
SEPARATOR := ";"
REJECTFILE := "rejected.csv"
EMPTY: string = ""
// ERROR MESSAGES
NOTVALIDSTATION :: "ERROR: No valid station selected as argument."
