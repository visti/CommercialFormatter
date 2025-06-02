package komm_fmt

import libc "core:c/libc/"
import "core:fmt"
import "core:os"
import filepath "core:path/filepath"
import str "core:strings"
import "core:time"

read_file :: proc(files: []string, currentStation: station) -> string {
	joinedFiles: string

	for file in files {
		baseFilename := filepath.stem(file)
		data, ok := os.read_entire_file(file, context.allocator)

		if !ok {
			panic("Could not read file.")
		}

		it := string(data)
		defer delete(it, context.allocator)

		if currentStation.hasHeadlines {
			lines := str.split_lines(string(data))
			it = str.join(lines[1:], "\n")
		}

		if currentStation.name == "Globus" {
			lines := str.split_lines(string(data))

			for &line, i in lines {
				line, _ = str.replace_all(line, " - ", ";")
				a := []string{baseFilename, line}
				if line != "" {
					lines[i] = str.join(a, "")
				}
				it = str.join(lines, "\n")
			}
			delete(lines, context.allocator)
		}
		clonedString := str.clone(it)
		a := []string{joinedFiles, clonedString}
		tempConcat := str.concatenate(a)
		joinedFiles = str.clone(tempConcat)
		delete(clonedString)
		delete(tempConcat)


	}
	return joinedFiles
}


clean_files :: proc(rejection_file: os.Handle, outputfile: os.Handle) {
	a := []os.Handle{rejection_file, outputfile}
	for file in a {
		path, _ := os.absolute_path_from_handle(file)
		slicePath := str.split(path, "/", context.allocator)
		db(string, path)
		size, err := os.file_size(file)
		if err == nil {
			if size == 0 {
				os.remove(path)
				fmt.printf("Removed empty file: %v\n", slicePath[len(slicePath) - 1])
			}
		}}

	outputfile_path, _ := os.absolute_path_from_handle(outputfile)
	b := []string{DELETE_COLS, outputfile_path}
	c := str.join(b, " ")
	command := fmt.caprintfln(c)
	defer delete(c)
	libc.system(command)
}

generate_rejection_filename :: proc(currentStation: station) -> string {
	currentTime, ok := time.time_to_datetime(NOW)

	if !ok {
		panic("could not get current time.")
	}

	datestring := fmt.tprintf("%d-%d-%d", currentTime.year, currentTime.month, currentTime.day)
	tempPath := str.join([]string{datestring, "reject", currentStation.name}, "-")
	REJECTFILE := tempPath
	defer delete(tempPath)
	rejectPath := filepath.join({REJECTDIR, REJECTFILE})
	return str.join({rejectPath, "csv"}, ".")
}
ask_user_stationtype :: proc() -> station {
	choice := str.to_upper(os.args[1])

	for s in Stations {
		if str.to_upper(s.name) == choice {
			info("Station: ", choice)
			return s
		}
	}
	fmt.printf("%v\n", NOTVALIDSTATION)
	os.exit(1)
}

ask_output_filename :: proc(currentStation: ^station) {
	buf: [256]byte
	fmt.println("Set output filename: ")
	n, err := os.read(os.stdin, buf[:])

	if err != nil {
		fmt.eprintln("Error: ", err)
	}

	if string(buf[:n]) == "" {
		fmt.eprintln("ERROR: Output filename cannot be empty!")
		os.exit(1)
	}

	currentStation.filename = str.clone(str.trim_space(string(buf[:n])))
}

get_files :: proc(currentStation: station) -> []string {
	ext := currentStation.ext
	filenames: [dynamic]string
	cwd := os.get_current_directory()
	f, err := os.open(cwd)
	defer os.close(f)
	files, _ := os.read_dir(f, 1)

	for fi in files {
		if fi.name != currentStation.filename {
			for x in ext {
				lowercaseFilename := str.to_lower(fi.name)
				if str.contains(lowercaseFilename, x) {
					append(&filenames, fi.name)

				}
			}}
	}
	if len(filenames) == 0 {
		info("", "No eligible files found.")
		os.exit(1)
	}

	foundFiles := str.join(filenames[:], ";")
	info("Found Files: ", foundFiles)
	delete(foundFiles)
	return filenames[:]
}

write_headlines :: proc(currentStation: station, file: os.Handle) {
	tempHeadlines := str.join(currentStation.headlines, ";")
	headlinesJoined := tempHeadlines
	defer delete(tempHeadlines)
	a := []string{headlinesJoined, "\n"}
	b := str.concatenate(a)
	os.write_string(file, b)
	delete(b)
}
