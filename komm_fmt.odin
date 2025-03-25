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
CONVERT :: "/mnt/d/komm_fmt/lib/convert.py"
REJECTDIR :: "/mnt/c/Users/eva/Gramex/Rapporteringer - Documents/Afviste_linjer_kom_land"
SEPARATOR := ";"
REJECTFILE := "rejected.csv"
DEFAULT_HEADLINES :: []string {
	"Date of Broadcasting",
	"Track starting time",
	"Track playing time",
	"Local-ID",
	"Track Title",
	"Main Artist",
	"Record Label",
	"GramexID",
	"Side",
	"Tracknummer",
	"Country of Recording",
	"Year of first release",
	"ISRC-Code",
}

// ERROR MESSAGES
NOTVALIDSTATION :: "ERROR: No valid station selected as argument."

station :: struct {
	name:         string,
	filename:     string,
	positions:    []int,
	stopwords:    []string,
	ext:          []string,
	hasHeadlines: bool,
	convert:      bool,
	positional:   bool,
	headlines:    []string,
}

Stations := []station{Bauer, Jyskfynske, Globus, Radio4}

Radio4 := station {
	name         = "Radio4",
	convert      = true,
	positional   = false,
	filename     = "",
	ext          = {"csv"},
	hasHeadlines = true,
	stopwords    = {},
	headlines    = {
		"DELETE",
		"Date of Broadcasting",
		"Track Starting Time",
		"DELETE",
		"Track Title",
		"Main Artist",
		"Track Playing Time",
		"Record Label",
		"Catalogue No",
		"DELETE",
		"DELETE",
		"DELETE",
	},
}


Globus := station {
	positional   = true,
	name         = "Globus",
	filename     = "",
	ext          = {"txt"},
	hasHeadlines = false,
	stopwords    = {
		"rasmus sk�tt",
		"dj pool",
		"FutureRecords",
		"dj deep",
		"megamwx",
		"mega mix",
		"yearmix",
		"year mix",
		"live mix",
		"mixtape",
		"mega",
		"summermix",
		"dancemix",
		"skøtt",
		"hit mix",
		"vi elsker",
		"pop mix",
		"dance mix",
		"mix cast",
		"mastermix",
		"retro mix",
		"summerparty",
		"dancemix",
		"in the mix",
		"mashup",
		"weekendmix",
		"dj cosmo",
		"dj simonsen",
		"martin deejay",
		"dj swa",
		"dj daniel olsen",
		"PHILIZZ",
		"maassive",
		"dj kosta",
		"dj tedu",
	},
	headlines    = {
		"Date of Broadcasting",
		"Track starting time",
		"Track playing time",
		"Main Artist",
		"Track Title",
	},
	positions    = {15, 13},
}

Jyskfynske := station {
	positional   = true,
	name         = "Jyskfynske",
	ext          = {"txt", "den"},
	hasHeadlines = false,
	positions    = {297, 291, 288, 267, 206, 200, 149, 98, 47, 26, 20, 11},
	headlines    = {
		"Date of Broadcasting", // 11
		"Track starting time", // 20
		"Track playing time", // 26
		"DELETE",
		"Album Title", // 47
		"Track Title", // 98
		"Main Artist", // 149
		"DELETE",
		"Record Label", // 200
		"Catalogue No.", // 206
		"Country of Recording", // 267
		"Year of First Release", // 288
		"DELETE", // 291
	},
	stopwords    = {
		"vejr",
		"vejle",
		";classic ",
		";dph ",
		"classic fm",
		";rw*",
		" rv ",
		" rek ",
		" intro ",
		"gmmj",
		" acc happy",
		"sponsor",
		"viborg",
		"rv_",
		"*rv",
		"wb_",
		" sw ",
		"sw_",
		"sweeper",
		"vlr",
		"toh ",
		"sw_unknown artist",
		"skala listen",
		"skala_",
		"_skala",
		"jfm",
		"festudvalget",
		"skala fm",
		"promo",
		"nyheder",
		"reklame",
		"toh skala",
		"fest_",
		"vejrsyd",
		"vejr [",
		"listen_count",
		"dst_",
		"fa_",
		"vib_",
		"rek-",
		"vi elsker - ",
		"happy hour",
		"trackbed",
	},
}

Bauer := station {
	positional   = true,
	name         = "Bauer",
	filename     = "",
	positions    = {185, 179, 173, 168, 163, 153, 128, 78, 28, 19, 14, 6},
	hasHeadlines = false,
	ext          = {"txt"},
	stopwords    = {
		"R1 ",
		" R1 ",
		"TOP HOUR",
		"99999998",
		"VO *:",
		"jingle",
		"SWEEPER",
		";SOFT ",
		"RADIO SOFT",
		"MR SW",
		";MR",
		"BREAKER",
		"NYHEDER",
		"PROMO",
		"VEJR",
		"Bauer",
	},
	headlines    = DEFAULT_HEADLINES,
}

printstation :: proc(station: station) {
	fmt.printf(
		"Name: %v\nPositions: %v\nFilename:%v\nStopwords:\n",
		station.name,
		station.positions,
		station.filename,
	)

	for word in station.stopwords {
		fmt.printf("%v\n", word)
	}
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
					append(&filenames, str.clone(fi.name))
				}
			}}
	}
	if len(filenames) == 0 {
		info("", "No eligible files found.")
		os.exit(1)
	}

	info("Found Files: ", str.join(filenames[:], ";"))
	return filenames[:]
}

write_headlines :: proc(currentStation: station, file: os.Handle) {
	headlinesJoined := str.join(currentStation.headlines, ";")
	a := []string{headlinesJoined, "\n"}
	os.write_string(file, str.concatenate(a))

}

info :: proc(field: string, string: string) {
	fmt.println("---------------------")
	fmt.printf("%v: %v\n", field, string)
	fmt.println("---------------------\n")
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

// Main entry point.
main :: proc() {
	EMPTY: string = ""

	stationChoice := ask_user_stationtype()

	if stationChoice.convert {
		command := fmt.caprint("/bin/python3 /mnt/d/komm_fmt/lib/convert.py *xls*")
		libc.system(command)
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

			for &line, i in lines {
				it = str.join(lines[1:], "\n")
			}
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
		a := []string{joinedFiles, str.clone(it)}
		joinedFiles = str.concatenate(a)
	}
	return joinedFiles
}

db :: proc($T: typeid, value: T) {
	if DEBUG {
		fmt.printf("DEBUG: %v\n", value)
	}
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
}

generate_rejection_filename :: proc(currentStation: station) -> string {
	currentTime, ok := time.time_to_datetime(NOW)

	if !ok {
		panic("could not get current time.")
	}

	datestring := fmt.tprintf("%d-%d-%d", currentTime.year, currentTime.month, currentTime.day)
	REJECTFILE := str.join([]string{datestring, "reject", currentStation.name}, "-")
	rejectPath := filepath.join({REJECTDIR, REJECTFILE})
	return str.join({rejectPath, "csv"}, ".")
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
