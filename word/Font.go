package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

var CLSID_Font = syscall.GUID{0x000209F5, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Font struct {
	Font_
}

func NewFont(pDisp *win32.IDispatch, addRef bool, scoped bool) *Font {
	p := &Font{Font_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewFontFromVar(v ole.Variant, addRef bool, scoped bool) *Font {
	return NewFont(v.PdispValVal(), addRef, scoped)
}

func NewFontInstance(scoped bool) (*Font, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_Font, nil, 
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_Font_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewFont(p, false, scoped), nil
}

