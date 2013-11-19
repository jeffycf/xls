package xls

/*
#cgo CFLAGS: -I/usr/local/libxls/include
#cgo LDFLAGS: -L/usr/local/libxls/lib -lxlsreader -liconv
#include "goxls.h"
*/
import "C"

func ReadSheet(filename string, sheetname string) string {
	var fname *C.char
	var shname *C.char
	fname = C.CString(filename)
	shname = C.CString(sheetname)
	C.readSheet(fname, shname)
	lines := C.GoString(C.lines)
	return lines
}
