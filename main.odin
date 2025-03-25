package komm_fmt

import libc "core:c/libc/"
import "core:fmt"
import io "core:io"
import "core:log"
import "core:mem"
import "core:os"
import filepath "core:path/filepath"
import str "core:strings"
import "core:time"
import dt "core:time/datetime"
import "core:unicode/utf8"

DEBUG :: false

NOW := time.now()
CONVERT :: "/bin/python3 /mnt/d/komm_fmt/lib/convert.py *xls*"
DELETE_COLS :: "/bin/python3 /mnt/d/komm_fmt/lib/delete_columns.py "
REJECTDIR :: "/mnt/c/Users/eva/Gramex/Rapporteringer - Documents/Afviste_linjer_kom_land"
SEPARATOR := ";"
REJECTFILE := "rejected.csv"
// ERROR MESSAGES
NOTVALIDSTATION :: "ERROR: No valid station selected as argument."


info :: proc(field: string, string: string) {
	fmt.println("---------------------")
	fmt.printf("%v: %v\n", field, string)
	fmt.println("---------------------\n")
}


// Main entry point.
main :: proc() {
	EMPTY: string = ""

	stationChoice := ask_user_stationtype()

	if stationChoice.convert {
		libc.system(CONVERT)
	}


	ask_output_filename(&stationChoice)
	defer delete(stationChoice.filename)
	outputFile := stationChoice.filename

	if outputFile == "" {
		fmt.eprintln("ERROR: Filename empty in main function.")
		os.exit(1)
	}

	if os.is_file(outputFile) {
		err := os.remove(outputFile)

		if err == nil {
			db(string, "Initialized output file.")
		}}

	os.write_entire_file(outputFile, transmute([]byte)(EMPTY))

	if len(os.args) == 1 {
		fmt.printf("Provide at least one argument.\n")
		fmt.printf("args: %v\n")
		for station in Stations {
			fmt.printf("Implemented types: %v\n", station.name)
		}

		return
	}
	filelist := get_files(stationChoice)
	newFile := read_file(filelist, stationChoice)
	defer delete(filelist)
	process_files(&newFile, stationChoice, outputFile)
}


db :: proc($T: typeid, value: T) {
	if DEBUG {
		fmt.printf("DEBUG: %v\n", value)
	}
}


check_for_stopwords :: proc(file: os.Handle, line: string, currentStation: station) -> string {
	lowercaseLine := str.to_lower(line) // Convert entire line to lowercase

	for word in currentStation.stopwords {
		if str.contains(lowercaseLine, str.to_lower(word)) { 	// Compare in lowercase
			a := []string{line, "\n"} // Write original line to rejected file
			os.write_string(file, str.concatenate(a))
			return "REJECT" // Stop early if a stopword is found
		}
	}
	return line // Only return the line if NO stopwords matched
}

wrap_up :: proc(rejected: int, processed: int) {
	fmt.printf("Processed Lines: %v\n", processed)
	fmt.printf("Rejected Lines: %v\n", rejected)
}
