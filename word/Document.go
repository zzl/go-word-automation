package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

var CLSID_Document = syscall.GUID{0x00020906, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Document struct {
	Document_
}

func NewDocument(pDisp *win32.IDispatch, addRef bool, scoped bool) *Document {
	 if pDisp == nil {
		return nil;
	}
	p := &Document{Document_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewDocumentFromVar(v ole.Variant, addRef bool, scoped bool) *Document {
	return NewDocument(v.IDispatch(), addRef, scoped)
}

func NewDocumentInstance(scoped bool) (*Document, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_Document, nil, 
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_Document_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewDocument(p, false, scoped), nil
}

func (this *Document) RegisterEventHandlers(handlers DocumentEvents2Handlers) uint32 {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_DocumentEvents2, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	dispImpl := &DocumentEvents2DispImpl{Handlers: handlers}
	disp := NewDocumentEvents2ComObj(dispImpl, false)
	
	var cookie uint32
	hr = cp.Advise(disp.IUnknown(), &cookie)
	win32.ASSERT_SUCCEEDED(hr)

	disp.Release()
	cp.Release()
	cpc.Release()
	return cookie
}

func (this *Document) UnRegisterEventHandlers(cookie uint32) {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_DocumentEvents2, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	hr = cp.Unadvise(cookie)
	win32.ASSERT_SUCCEEDED(hr)

	cp.Release()
	cpc.Release()
}

