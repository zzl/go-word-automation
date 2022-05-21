package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

var CLSID_ParagraphFormat = syscall.GUID{0x000209F4, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ParagraphFormat struct {
	ParagraphFormat_
}

func NewParagraphFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ParagraphFormat {
	p := &ParagraphFormat{ParagraphFormat_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewParagraphFormatFromVar(v ole.Variant, addRef bool, scoped bool) *ParagraphFormat {
	return NewParagraphFormat(v.PdispValVal(), addRef, scoped)
}

func NewParagraphFormatInstance(scoped bool) (*ParagraphFormat, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_ParagraphFormat, nil, 
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_ParagraphFormat_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewParagraphFormat(p, false, scoped), nil
}

