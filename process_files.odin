package komm_fmt

import "core:fmt"
import "core:os"
import str "core:strings"
import "core:time"

check_for_stopwords :: proc(file: os.Handle, line: string, currentStation: station) -> string {
	lowercaseLine := str.to_lower(line) // Convert entire line to lowercase

	for word in currentStation.stopwords {
		if str.contains(lowercaseLine, str.to_lower(word)) { 	// Compare in lowercase
			a := []string{line, "\n"} // Write original line to rejected file
			b := str.concatenate(a)
			os.write_string(file, b)
			delete(b)
			return str.clone("REJECT") // Stop early if a stopword is found
		}
	}
	return str.clone(line) // Only return the line if NO stopwords matched
}

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
	rejCreateOK := os.write_entire_file_or_err(rejectPath, transmute([]byte)(EMPTY))
	if rejCreateOK != nil {
		fmt.eprintln("ERROR: Failed to write rejection file:", rejCreateOK)
		os.exit(1)
	}

	if !os.is_file(rejectPath) {
		fmt.println(rejectPath)
		fmt.eprintln("ERROR: File was not created:", rejectPath)
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
				temp := str.join(a, ";")
				delete(checkedLine)
				checkedLine = temp
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
			} else {
				clonedString := str.clone(checkedLine)
				outputLine = str.join([]string{clonedString, "\n"}, "")
				delete(clonedString)
			}

			_, err := os.write_string(outputFileHandle, outputLine)
			delete(outputLine)
		}

		delete(checkedLine)

		if err != nil {
			fmt.printf("%v\n", err)
		}
		delete(parts)
	}
	wrap_up(rejected, processed)

	clean_files(rejectionFile, outputFileHandle)
}
