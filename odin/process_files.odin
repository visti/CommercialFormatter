package komm_fmt

import "core:fmt"
import "core:os"
import filepath "core:path/filepath"
import str "core:strings"

check_for_stopwords :: proc(
    file: os.Handle,
    line: string,
    currentStation: station,
    stopwords: [dynamic]string,
) -> string {
    lowercaseLine := str.to_lower(line)

    for word in stopwords {
        if str.contains(lowercaseLine, str.to_lower(word)) {
            a := []string{line, "\n"}
            b := str.concatenate(a)
            os.write_string(file, b)
            delete(b)
            return str.clone("REJECT")
        }
    }
    return str.clone(line)
}

// Build "<output>_additional<ext>" next to the main output
make_additional_filename :: proc(base: string) -> string {
    ext := filepath.ext(base)
    stem := filepath.stem(base)
    dir := filepath.dir(base)
    with_postfix := str.join([]string{stem, ADDITIONAL_POSTFIX, ext}, "")
    if dir != "" {
        return str.join([]string{dir, "/", with_postfix}, "")
    }
    return with_postfix
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
    checkedLine: string

    // Additional routing controls
    additionalFileHandle: os.Handle
    additionalPath: string
    hasAdditional := (len(ADDITIONAL_FILTER) > 0)

    // Keep existing behavior for separator
    if currentStation.name == "Globus" {
        SEPARATOR = ":"
    } else {
        SEPARATOR = ";"
    }

    // Prepare output file
    outputFileHandle, err := os.open(outputFile, os.O_CREATE | os.O_WRONLY | os.O_TRUNC)
    if err != nil {
        fmt.eprintln("ERROR: could not create output file:", err)
    }
    defer os.close(outputFileHandle)

    // Create rejection file (empty) and open handle
    rejectPath := generate_rejection_filename(currentStation)
    rejCreateOK := os.write_entire_file_or_err(rejectPath, transmute([]byte)(EMPTY))
    if rejCreateOK != nil {
        fmt.eprintln("ERROR: Failed to write rejection file:", rejCreateOK)
        os.exit(1)
    }
    if !os.is_file(rejectPath) {
        fmt.eprintln("ERROR: File was not created:", rejectPath)
        os.exit(1)
    }
    rejectionFile, rejectfile_open_error := os.open(rejectPath, 2)
    defer os.close(rejectionFile)
    if rejectfile_open_error != nil {
        fmt.eprintln("Error creating rejectfile: ", rejectfile_open_error)
        os.exit(1)
    }

    // Open additional file if enabled (no handle shadowing)
    if hasAdditional {
        additionalPath = make_additional_filename(outputFile)

        h, addErr := os.open(additionalPath, os.O_CREATE | os.O_WRONLY | os.O_TRUNC)
        if addErr != nil {
            fmt.eprintln("ERROR: could not create additional file:", addErr)
            os.exit(1)
        }
        additionalFileHandle = h

        // Info so we can see target path & confirm handle
        fmt.println("Additional file opened for write:", additionalPath)

        // Same headlines as standard output
        write_headlines(currentStation, additionalFileHandle)
    }

    // Write headlines to output and rejection (preserve current behavior)
    write_headlines(currentStation, outputFileHandle)
    write_headlines(currentStation, rejectionFile)

    // Build stopword list: DEFAULT + station-specific
    stops: [dynamic]string
    for word in DEFAULT_STOPWORDS {
        append(&stops, word)
    }
    for word in currentStation.stopwords {
        append(&stops, word)
    }

    // Ensure positional cut indices are in ascending order for safe left-to-right slicing
    if currentStation.positional {
        for i := 0; i < len(currentStation.positions); i += 1 {
            for j := i + 1; j < len(currentStation.positions); j += 1 {
                if currentStation.positions[j] < currentStation.positions[i] {
                    tmp := currentStation.positions[i]
                    currentStation.positions[i] = currentStation.positions[j]
                    currentStation.positions[j] = tmp
                }
            }
        }
    }

    for line in str.split_lines_iterator(file) {
        parts: [dynamic]string

        if USE_STOPWORDS {
            checkedLine = check_for_stopwords(rejectionFile, line, currentStation, stops)
        } else {
            checkedLine = str.clone(line)
        }

        if checkedLine != "REJECT" {
            // After stopword check: handle additional routing (case-insensitive)
            routeToAdditional := false
            if hasAdditional {
                cl := str.to_lower(checkedLine)
                af := str.to_lower(ADDITIONAL_FILTER)
                routeToAdditional = str.contains(cl, af)
            }

            if currentStation.positional {
                start := 0

                // Slice left-to-right using sorted positions
                for pos in currentStation.positions {
                    lineLength := len(checkedLine)
                    effective_pos := pos
                    if effective_pos > lineLength {
                        effective_pos = lineLength
                    }

                    if effective_pos < start {
                        // Safeguard: skip invalid (out-of-order) cut points
                        continue
                    }

                    part := (string)(checkedLine)[start:effective_pos]
                    trimmedPart := str.trim_space(part)
                    append(&parts, trimmedPart)
                    start = effective_pos
                }
            }

            processed += 1
            outputLine: string

            if currentStation.positional {
                modifiedLine := str.join(parts[:], SEPARATOR)
                outputLine = str.join([]string{modifiedLine, "\n"}, "")
                delete(modifiedLine)
			} else {
    fields := str.split(checkedLine, ";")
    if len(fields) > 1 && len(fields[1]) == 6 {
        rawTime := fields[1]
        hours := rawTime[0:2]
        minutes := rawTime[2:4]
        seconds := rawTime[4:6]
        formattedTime := str.join([]string{hours, ":", minutes, ":", seconds}, "")
        fields[1] = formattedTime
    }

    modifiedLine := str.join(fields[:], ";")
    outputLine = str.join([]string{modifiedLine, "\n"}, "")
    delete(modifiedLine)
}





            if routeToAdditional {
                if hasAdditional {
                    _, werr := os.write_string(additionalFileHandle, outputLine)
                    if werr != nil {
                        fmt.printf("WRITE additional error: %v\n", werr)
                    }
                }
            } else {
                _, werr := os.write_string(outputFileHandle, outputLine)
                if werr != nil {
                    fmt.printf("WRITE output error: %v\n", werr)
                }
            }

            delete(outputLine)
            delete(parts)
        } else {
            rejected += 1
        }

        delete(checkedLine)
    }
    wrap_up(rejected, processed)

    // --- CLEANUP PHASE ---

    if hasAdditional {
        os.close(additionalFileHandle)

        // Check if the additional file exists
        exists, ferr := os.stat(additionalPath)
        if ferr == nil {
            contents, ok := os.read_entire_file(additionalPath)
            if ok {
                // Convert []u8 → string for counting
                text := string(contents)
                line_count := str.count(text, "\n")
                if line_count <= 1 {
                    del_err := os.remove(additionalPath)
                    if del_err == nil {
                        fmt.printf("Deleted empty additional file: %s\n", additionalPath)
                    } else {
                        fmt.printf("Failed to delete %s: %v\n", additionalPath, del_err)
                    }
                }
                delete(text)
            } else {
                fmt.printf("Failed to read %s: %v\n", additionalPath)
            }
        }
        _ = exists // silence unused var if needed
    }

    clean_files(rejectionFile, outputFileHandle)
}
