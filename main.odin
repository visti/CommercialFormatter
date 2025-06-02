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
	when ODIN_DEBUG {
		delete_existing_file(MEMORY_LOG)
		// Open log file in append mode (or use O_TRUNC if you prefer to overwrite)
		logFile, err := os.open(MEMORY_LOG, os.O_CREATE | os.O_WRONLY | os.O_APPEND)
		if err != nil {
			// Fallback: if opening the file fails, print an error to stderr.
			fmt.eprintf("Error opening memory leak log file: %v\n", err)
		}
		defer os.close(logFile)

		track: mem.Tracking_Allocator
		mem.tracking_allocator_init(&track, context.allocator)
		context.allocator = mem.tracking_allocator(&track)

		defer {
			if len(track.allocation_map) > 0 {
				for _, entry in track.allocation_map {
					// Write leak info to the log file instead of printing
					fmt.fprintf(logFile, "%v leaked %v bytes\n", entry.location, entry.size)
				}
			}
			if len(track.bad_free_array) > 0 {
				for entry in track.bad_free_array {
					fmt.fprintf(logFile, "%v bad free at %v\n", entry.location, entry.memory)
				}
			}
			mem.tracking_allocator_destroy(&track)
		} 
	}
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


	delete_existing_file(outputFile)
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
