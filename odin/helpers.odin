package komm_fmt

import "core:fmt"
import "core:os"
info :: proc(field: string, string: string) {
	fmt.println("---------------------")
	fmt.printf("%v: %v\n", field, string)
	fmt.println("---------------------\n")
}


db :: proc($T: typeid, value: T) {
	if DEBUG {
		fmt.printf("DEBUG: %v\n", value)
	}
}

delete_existing_file :: proc(file: string) {

	if os.is_file(file) {
		err := os.remove(file)

		if err == nil {
			fmt.println("Initialized", file)
		}}
}
