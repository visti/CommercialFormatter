package komm_fmt

import "core:fmt"

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
