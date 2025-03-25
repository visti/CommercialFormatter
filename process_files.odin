package komm_fmt

import "core:fmt"
import "core:os"
import str "core:strings"
import "core:time"

process_files :: proc(
	file: ^string,
	currentStation: station,
	outputFile: string,
	loc := #caller_location,
) {
	modifiedLine: [dynamic]string
	headlinesOut: string
	EMPTY: string
	rejected: int
	processed: int
	outputLine: string
	currenttime, ok := time.time_to_datetime(NOW)
	rejectPath := generate_rejection_filename(currentStation)
	defer delete(rejectPath)

	fmt.println("Rejection file saved to: ", rejectPath)
	if currentStation.name == "Globus" {SEPARATOR = ":"}

	//create rejection file
	os.write_entire_file(rejectPath, transmute([]byte)(EMPTY))

	if !os.is_file(rejectPath) {
		fmt.eprintln("ERROR: File was not created:", outputFile)
		os.exit(1)
	}
	rejectionFile, rejectfile_open_error := os.open(rejectPath, 2)
	defer os.close(rejectionFile)

	if rejectfile_open_error != nil {
		fmt.eprintln("Error creating rejectfile: ", rejectfile_open_error)
	}

	outputFileHandle, err := os.open(outputFile, os.O_CREATE | os.O_WRONLY | os.O_TRUNC)
	if err != nil {
		fmt.eprintln("ERROR: could not create output file:", err)
	}
	defer os.close(outputFileHandle)

	write_headlines(currentStation, outputFileHandle)
	write_headlines(currentStation, rejectionFile)

	if err != nil {
		fmt.printf("%v\n", err)
	}

	for line in str.split_lines_iterator(file) {
		parts: [dynamic]string
		start: int

		checkedLine := check_for_stopwords(rejectionFile, line, currentStation)

		if checkedLine != "REJECT" {
			if currentStation.name == "Radio4" {
				a := []string{checkedLine[:17], checkedLine[17:]}
				checkedLine = str.join(a, ";")
			}

			if currentStation.positional {
				#reverse for &position in currentStation.positions {
					lineLength := len(checkedLine)
					if position > lineLength {
						position = lineLength // Prevent out-of-bounds slicing
					}
					part := (string)(checkedLine)[start:position]

					trimmedPart := str.trim_space(part)
					append(&parts, trimmedPart)

					start = position
				}
			}
			processed += 1
		} else {rejected += 1}

		// construct output line from parts
		if checkedLine != "REJECT" {
			if currentStation.positional {
				append(&parts, checkedLine[start:])
				modifiedLine := str.join(parts[:], SEPARATOR)

				outputLine = str.join([]string{modifiedLine, "\n"}, "")
				delete(modifiedLine)
			} else {outputLine = str.join([]string{str.clone(checkedLine), "\n"}, "")}

			_, err := os.write_string(outputFileHandle, outputLine)
			delete(outputLine)
		}

		if err != nil {
			fmt.printf("%v\n", err)
		}

		delete(parts)
	}
	wrap_up(rejected, processed)

	clean_files(rejectionFile, outputFileHandle)
}
