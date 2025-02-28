
package main

import "core:fmt"
import "core:os"

// Declare a mutable global variable.

assign_variable :: proc(VARI: string) {
    buf: [256]byte
    fmt.println("Set output filename:")
    
    n, err := os.read(os.stdin, buf[:])
    if err != nil {
        fmt.eprintln("Error: ", err)
        return
    }
    
    // Convert input to a string.
    input_str := string(buf[:n])
    // Debug print to check what we got:
    fmt.printf("Raw input: %q\n", input_str)
    
    // If the input string is empty (or just whitespace), warn and exit.
    if len(input_str) == 0 {
        fmt.println("No input provided.")
        return
    }
    
    // Inline new_clone strategy: new_clone allocates a heap copy.
    VARIABLE = new_clone(input_str)^
    fmt.printf("Inside assign_variable, VARIABLE: %v\n", VARIABLE)

}

main :: proc() {
VARIABLE: string = "test.csv"

    fmt.printf("Before assign_variable, VARIABLE: %v\n", VARIABLE)
    assign_variable(VARIABLE)
    fmt.printf("After assign_variable, VARIABLE: %v\n", VARIABLE)
}
