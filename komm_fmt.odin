package komm_fmt

import "core:fmt"
import io "core:io"
import "core:os"
import str "core:strings"
import "core:unicode/utf8"

DEBUG :: true

// ERROR MESSAGES
NOT_VALID_STATION :: "ERROR: No valid station selected as argument."

Station :: struct {
	name:          string,
	filename:      string,
	positions:     []int,
	stopwords:     []string,
	ext:           []string,
	has_headlines: bool,
	headlines:     []string,
}

Stations := []Station{Bauer, Jyskfynske}

Jyskfynske := Station {
	name          = "Jyskfynske",
	ext           = {"txt", "den"},
	has_headlines = false,
	positions     = {297, 291, 288, 267, 216, 200, 149, 98, 47, 37, 28, 27, 20, 11},
}

Bauer := Station {
	name          = "Bauer",
	positions     = {185, 179, 173, 168, 163, 153, 128, 78, 28, 19, 14, 6},
	has_headlines = false,
	ext           = {"txt"},
	stopwords     = {
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
	headlines     = {
		"Date of Broadcasting",
		"Track starting time",
		"Track playing time",
		"Local-ID;Track Title",
		"Main Artist",
		"Record Label",
		"GramexID",
		"Side",
		"Tracknummer",
		"Country of Recording",
		"Year of first release",
		"ISRC-Code",
	},
}

printstation :: proc(station: Station) {
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

get_files :: proc(ext: ''string) -> []string {
	output_files: string
	a: []string

	cwd := os.get_current_directory()
	f, err := os.open(cwd)
	defer os.close(f)

	files, _ := os.read_dir(f, 1)

	for fi in files {
		for x in ext {
			lowercase_filename := str.to_lower(fi.name)
			if str.contains(lowercase_filename, x) {
				a := []string{output_files, fi.name}
				output_files = str.join(a, ";")
			}
		}
	}
	db(string, "output of get_files:")
	db([]string, str.split(output_files[1:], ";"))
	return str.split(output_files[1:], ";")
}

ask_user_stationtype :: proc() -> Station {
	choice := str.to_upper(os.args[1])
	switch choice {
	case "BAUER":
		fmt.printf("%v\n", "Bauer")
		return Bauer
	case "JYSKFYNSKE":
		fmt.printf("%v\n", "Jysk")
		return Jyskfynske
	case:
		fmt.printf("%v\n", NOT_VALID_STATION)
		os.exit(1)

	}
	return Station{}

}

// Main entry point.
main :: proc() {
	if len(os.args) == 1 {
		fmt.printf("Provide at least one argument.\n")
		fmt.printf("args: %v\n")
		for station in Stations {
			fmt.printf("Implemented types: %v\n", station.name)
		}

		return
	}


	stationChoice := ask_user_stationtype()

	filelist := get_files(stationChoice.ext)
	new_file := read_file(filelist)
	process_files(&new_file, stationChoice)
}


read_file :: proc(files: []string) -> string {
	joined_files: string
	imported_files: string
	for file in files {
		data, ok := os.read_entire_file(file, context.allocator)

		if !ok {
			panic("could not read file")
		}
		it := string(data)

		a := []string{joined_files, it}
		joined_files = str.concatenate(a)

	}
	db(string, "Input lines:")
	// db(string, joined_files)
	return joined_files
}

db :: proc($T: typeid, value: T) {
	if DEBUG {
		fmt.printf("DEBUG: %v\n", value)
	}
}


process_files :: proc(file: ^string, station: Station) {
	db(string, "process_lines started.")
	modified_line: [dynamic]string
	processed_lines: [dynamic]string


	db(string, "positions:")
	db([]int, station.positions)


	for line in str.split_lines_iterator(file) {
		parts: [dynamic]string
		start: int

		#reverse for pos in station.positions {
			part := line[start:pos]
			trimmed_part := str.trim_space(part)
			append(&parts, trimmed_part)
			start = pos
		}

		append(&parts, line[start:])

		modified_line := str.join(parts[:], ";")
		append(&processed_lines, modified_line)

		final_output := str.join(processed_lines[:], "\n")
		os.write_entire_file("output.csv", transmute([]byte)(final_output))
		//output := str.join(processed_lines, "\n")
		//os.write_entire_file("output.txt", transmute([]byte)(processed_lines))
	}

	db([dynamic]string, processed_lines)
}
