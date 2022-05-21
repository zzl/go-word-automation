package main

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-word-automation/word"
	"log"
	"strconv"
)

func main() {

	ole.Initialize()
	defer ole.Uninitialize()

	//
	scope := com.NewScope()
	defer scope.Leave()

	//
	wApp, err := word.NewApplicationInstance(true)
	if err != nil {
		log.Fatal(err)
	}
	wApp.SetVisible(true)
	wApp.Documents().Add()

	closeCount := 3
	var cookie uint32
	cookie = wApp.RegisterEventHandlers(word.ApplicationEvents4Handlers{
		DocumentBeforeClose: func(doc *word.Document, cancel *win32.VARIANT_BOOL) {
			msg := "##before close #" + strconv.Itoa(closeCount) + "... "
			doc.Paragraphs().Add(doc.Content()).Range().InsertBefore(msg)
			*cancel = win32.VARIANT_TRUE
			closeCount -= 1
			if closeCount == 0 {
				wApp.UnRegisterEventHandlers(cookie)
				win32.PostQuitMessage(0)
			}
		},
	})

	var msg win32.MSG
	for {
		ret, _ := win32.GetMessage(&msg, 0, 0, 0)
		if ret == 0 {
			break
		}
		win32.DispatchMessage(&msg)
	}
}
