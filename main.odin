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

// Main entry point.
main :: proc() {

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

wrap_up :: proc(rejected: int, processed: int) {
	fmt.printf("Processed Lines: %v\n", processed)
	fmt.printf("Rejected Lines: %v\n", rejected)
}
