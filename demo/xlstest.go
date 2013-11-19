package main

import (
	"fmt"
	"github.com/zhcy/xls"
)

func main() {
	lines := xls.ReadSheet("test2.xls", "Sheet1")
	fmt.Println(lines)
	fmt.Println("over")
}
