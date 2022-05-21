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

	wApp, err := word.NewApplicationInstance(true)
	if err != nil {
		log.Fatal(err)
	}
	wApp.SetVisible(true)

	filePath := ole.VarScoped(`C:\Windows\System32\license.rtf`)
	doc := wApp.Documents().Open(&filePath, ole.NamedArgs{
		"ConfirmConversions": false,
		"ReadOnly":           true,
		"AddToRecentFiles":   true,
		"Format":             word.WdOpenFormat.WdOpenFormatAuto,
	})
	rng := doc.Content()
	find := rng.Find()
	if find.Execute(ole.NamedArgs{
		"FindText":    "Microsoft",
		"MatchCase":   true,
		"ReplaceWith": "MICRO$OFT",
	}) {
		font := rng.Font()
		font.SetColorIndex(word.WdColorIndex.WdBlue)
		font.SetBold(1)
		font.SetItalic(1)
	}

}
