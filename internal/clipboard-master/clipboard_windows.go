// Copyright 2013 @atotto. All rights reserved.
// Use of this source code is governed by a BSD-style
// license that can be found in the LICENSE file.

//go:build windows
// +build windows

package clipboard

import (
	"encoding/binary"
	"runtime"
	"syscall"
	"time"
	"unsafe"
)

const (
	cfUnicodetext = 13
	gmemMoveable  = 0x0002
)

var (
	user32                     = syscall.MustLoadDLL("user32")
	isClipboardFormatAvailable = user32.MustFindProc("IsClipboardFormatAvailable")
	openClipboard              = user32.MustFindProc("OpenClipboard")
	closeClipboard             = user32.MustFindProc("CloseClipboard")
	emptyClipboard             = user32.MustFindProc("EmptyClipboard")
	getClipboardData           = user32.MustFindProc("GetClipboardData")
	setClipboardData           = user32.MustFindProc("SetClipboardData")

	kernel32     = syscall.NewLazyDLL("kernel32")
	globalAlloc  = kernel32.NewProc("GlobalAlloc")
	globalFree   = kernel32.NewProc("GlobalFree")
	globalLock   = kernel32.NewProc("GlobalLock")
	globalUnlock = kernel32.NewProc("GlobalUnlock")
	lstrcpy      = kernel32.NewProc("lstrcpyW")

	registerClipboardFormatA = user32.MustFindProc("RegisterClipboardFormatA")
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
	// LockOSThread ensure that the whole method will keep executing on the same thread from begin to end (it actually locks the goroutine thread attribution).
	// Otherwise if the goroutine switch thread during execution (which is a common practice), the OpenClipboard and CloseClipboard will happen on two different threads, and it will result in a clipboard deadlock.
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()
	if formatAvailable, _, err := isClipboardFormatAvailable.Call(cfUnicodetext); formatAvailable == 0 {
		return "", err
	}
	err := waitOpenClipboard()
	if err != nil {
		return "", err
	}

	h, _, err := getClipboardData.Call(cfUnicodetext)
	if h == 0 {
		_, _, _ = closeClipboard.Call()
		return "", err
	}

	l, _, err := globalLock.Call(h)
	if l == 0 {
		_, _, _ = closeClipboard.Call()
		return "", err
	}

	text := syscall.UTF16ToString((*[1 << 20]uint16)(unsafe.Pointer(l))[:])

	r, _, err := globalUnlock.Call(h)
	if r == 0 {
		_, _, _ = closeClipboard.Call()
		return "", err
	}

	closed, _, err := closeClipboard.Call()
	if closed == 0 {
		return "", err
	}
	return text, nil
}

func writeAll(text string) error {

	// LockOSThread ensure that the whole method will keep executing on the same thread from begin to end (it actually locks the goroutine thread attribution).
	// Otherwise if the goroutine switch thread during execution (which is a common practice), the OpenClipboard and CloseClipboard will happen on two different threads, and it will result in a clipboard deadlock.
	runtime.LockOSThread()
	defer runtime.UnlockOSThread()

	err := waitOpenClipboard()
	if err != nil {
		return err
	}

	r, _, err := emptyClipboard.Call(0)
	if r == 0 {
		_, _, _ = closeClipboard.Call()
		return err
	}

	// data, _ := syscall.UTF16FromString(text)
	//for FileMaker types, the first 4 bytes on the clipboard is the size of the data on the clipboard
	const offset = 4
	data := []byte(text)

	// "If the hMem parameter identifies a memory object, the object must have
	// been allocated using the function with the GMEM_MOVEABLE flag."
	hMem, _, err := globalAlloc.Call(gmemMoveable, uintptr(len(data)*int(unsafe.Sizeof(data[0]))+offset))

	if hMem == 0 {
		_, _, _ = closeClipboard.Call()
		return err
	}
	defer func() {
		if hMem != 0 {
			globalFree.Call(hMem)
		}
	}()

	l, _, err := globalLock.Call(hMem)
	if l == 0 {
		_, _, _ = closeClipboard.Call()
		return err
	}

	r, _, err = lstrcpy.Call(l+offset, uintptr(unsafe.Pointer(&data[0])))
	if r == 0 {
		_, _, _ = closeClipboard.Call()
		return err
	}

	size := len(text)
	sizeBytes := [4]byte{}
	binary.LittleEndian.PutUint32(sizeBytes[:], uint32(size))

	//2 null bytes in a row, basically
	//https://superuser.com/questions/946533/is-there-any-way-to-copy-null-bytes-ascii-0x00-to-the-clipboard-on-windows
	r, _, err = lstrcpy.Call(l, uintptr(unsafe.Pointer(&sizeBytes[0])))
	if r == 0 {
		return err
	}

	r, _, err = globalUnlock.Call(hMem)
	if r == 0 {
		if err.(syscall.Errno) != 0 {
			_, _, _ = closeClipboard.Call()
			return err
		}
	}

	format := "Mac-XMTB"
	formatBytes := append([]byte(format), 0)
	uFormat, _, _ := registerClipboardFormatA.Call(uintptr(unsafe.Pointer(&formatBytes[0])))

	r, _, err = setClipboardData.Call(uFormat, hMem)
	if r == 0 {
		_, _, _ = closeClipboard.Call()
		return err
	}
	hMem = 0 // suppress deferred cleanup
	closed, _, err := closeClipboard.Call()
	if closed == 0 {
		return err
	}
	return nil
}
