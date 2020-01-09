// Copyright 2013 @atotto. All rights reserved.
// Use of this source code is governed by a BSD-style
// license that can be found in the LICENSE file.

// +build windows

package clipboard

import (
	"syscall"
	"time"
	"unsafe"

	"encoding/binary" //20/01/01 Takeshi Yamamoto
)

const (
	cfUnicodetext = 13
	gmemMoveable  = 0x0002
)

var (
	user32           = syscall.MustLoadDLL("user32")
	openClipboard    = user32.MustFindProc("OpenClipboard")
	closeClipboard   = user32.MustFindProc("CloseClipboard")
	emptyClipboard   = user32.MustFindProc("EmptyClipboard")
	getClipboardData = user32.MustFindProc("GetClipboardData")
	setClipboardData = user32.MustFindProc("SetClipboardData")

	registerClipboardFormat = user32.MustFindProc("RegisterClipboardFormatA")
		
	kernel32     = syscall.NewLazyDLL("kernel32")
	globalAlloc  = kernel32.NewProc("GlobalAlloc")
	globalFree   = kernel32.NewProc("GlobalFree")
	globalLock   = kernel32.NewProc("GlobalLock")
	globalUnlock = kernel32.NewProc("GlobalUnlock")
	lstrcpy      = kernel32.NewProc("lstrcpyW")
)

// waitOpenClipboard opens the clipboard, waiting for up to a second to do so.
func waitOpenClipboard() error {
	started := time.Now()
	limit := started.Add(time.Second)
	var r uintptr
	var err error
	for time.Now().Before(limit) {
		r, _, err = openClipboard.Call(0)
		if r != 0 {
			return nil
		}
		time.Sleep(time.Millisecond)
	}
	return err
}

func readAll() (string, error) {
	err := waitOpenClipboard()
	if err != nil {
		return "", err
	}
	defer closeClipboard.Call()

	h, _, err := getClipboardData.Call(cfUnicodetext)
	if h == 0 {
		return "", err
	}

	l, _, err := globalLock.Call(h)
	if l == 0 {
		return "", err
	}

	text := syscall.UTF16ToString((*[1 << 20]uint16)(unsafe.Pointer(l))[:])

	r, _, err := globalUnlock.Call(h)
	if r == 0 {
		return "", err
	}

	return text, nil
}

func writeAll(text string) error {
	err := waitOpenClipboard()
	if err != nil {
		return err
	}
	defer closeClipboard.Call()

	r, _, err := emptyClipboard.Call(0)
	if r == 0 {
		return err
	}

	//for FileMaker types, the first 4 bytes on the clipboard is the size of the data on the clipboard
	const offset = 4

	//data := syscall.StringToUTF16(text)
	//https://qiita.com/ikawaha/items/3c3994559dfeffb9f8c9
	data := []byte(text)

	// "If the hMem parameter identifies a memory object, the object must have
	// been allocated using the function with the GMEM_MOVEABLE flag."
	// h, _, err := globalAlloc.Call(gmemMoveable, uintptr(len(data)*int(unsafe.Sizeof(data[0]))))
	h, _, err := globalAlloc.Call(gmemMoveable, uintptr(len(data)*int(unsafe.Sizeof(data[0])))+offset)

	if h == 0 {
		return err
	}
	defer func() {
		if h != 0 {
			globalFree.Call(h)
		}
	}()

	l, _, err := globalLock.Call(h)
	if l == 0 {
		return err
	}

	r, _, err = lstrcpy.Call(l+offset, uintptr(unsafe.Pointer(&data[0])))
	if r == 0 {
		return err
	}	

	/* 20/01/01 Takeshi Yamamoto ここから */
	//for FileMaker types, the first 4 bytes on the clipboard is the size of the data on the clipboard
	//https://wa3.i-3-i.info/word11428.html
	//https://kotaeta.com/54051254
	x := len(text)
    y := [4]byte{}
    binary.LittleEndian.PutUint32(y[:], uint32(x))

	//2 null bytes in a row, basically
	//https://superuser.com/questions/946533/is-there-any-way-to-copy-null-bytes-ascii-0x00-to-the-clipboard-on-windows
	r, _, err = lstrcpy.Call(l, uintptr(unsafe.Pointer(&y[0])))
	if r == 0 {
		return err
	}	
	/* 20/01/01 Takeshi Yamamoto ここまで */

	r, _, err = globalUnlock.Call(h)
	if r == 0 {
		if err.(syscall.Errno) != 0 {
			return err
		}
	}

	/* 20/01/01 Takeshi Yamamoto ここから */
	//https://stackoverflow.com/questions/51925111/passing-string-to-syscalluintptr/51925586
	formatString := "Mac-XMTB"
	formatByte := append([]byte(formatString), 0)
	formatId, _, _ := registerClipboardFormat.Call(uintptr(unsafe.Pointer(&formatByte[0])))

	//r, _, err = setClipboardData.Call(cfUnicodetext, h)
	r, _, err = setClipboardData.Call(formatId, h)
	/* 20/01/01 Takeshi Yamamoto ここまで */

	if r == 0 {
		return err
	}
	h = 0 // suppress deferred cleanup
	return nil
}
