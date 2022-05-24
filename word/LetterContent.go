package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

var CLSID_LetterContent = syscall.GUID{0x000209F1, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type LetterContent struct {
	LetterContent_
}

func NewLetterContent(pDisp *win32.IDispatch, addRef bool, scoped bool) *LetterContent {
	 if pDisp == nil {
		return nil;
	}
	p := &LetterContent{LetterContent_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewLetterContentFromVar(v ole.Variant, addRef bool, scoped bool) *LetterContent {
	return NewLetterContent(v.IDispatch(), addRef, scoped)
}

func NewLetterContentInstance(scoped bool) (*LetterContent, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_LetterContent, nil, 
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_LetterContent_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewLetterContent(p, false, scoped), nil
}

