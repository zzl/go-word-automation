package main

import (
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"github.com/zzl/go-word-automation/word"
	"log"
)

func main() {

	ole.Initialize()
	defer ole.Uninitialize()

	defer com.NewScope().Leave()
	//
	wApp, err := word.NewApplicationInstance(true)
	if err != nil {
		log.Fatal(err)
	}

	wApp.SetVisible(true)
	doc := wApp.Documents().Add()

	wApp.FontNames().ForEach(func(font string) bool {
		p := doc.Paragraphs().Add()
		p.Range().InsertAfter(font)
		p.Range().Font().SetName(font)
		return true
	})

	com.WithScope(func() {

	})
}
