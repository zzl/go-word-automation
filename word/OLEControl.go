package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

var CLSID_OLEControl = syscall.GUID{0x000209F2, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEControl struct {
	OLEControl_
}

func NewOLEControl(pDisp *win32.IDispatch, addRef bool, scoped bool) *OLEControl {
	 if pDisp == nil {
		return nil;
	}
	p := &OLEControl{OLEControl_{ole.OleClient{pDisp}}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func NewOLEControlFromVar(v ole.Variant, addRef bool, scoped bool) *OLEControl {
	return NewOLEControl(v.IDispatch(), addRef, scoped)
}

func NewOLEControlInstance(scoped bool) (*OLEControl, error) {
	var p *win32.IDispatch
	hr := win32.CoCreateInstance(&CLSID_OLEControl, nil, 
		win32.CLSCTX_INPROC_SERVER|win32.CLSCTX_LOCAL_SERVER,
		&IID_OLEControl_, unsafe.Pointer(&p))
	if win32.FAILED(hr) {
		return nil, com.NewError(hr)
	}
	return NewOLEControl(p, false, scoped), nil
}

func (this *OLEControl) RegisterEventHandlers(handlers OCXEventsHandlers) uint32 {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_OCXEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	dispImpl := &OCXEventsDispImpl{Handlers: handlers}
	disp := NewOCXEventsComObj(dispImpl, false)
	
	var cookie uint32
	hr = cp.Advise(disp.IUnknown(), &cookie)
	win32.ASSERT_SUCCEEDED(hr)

	disp.Release()
	cp.Release()
	cpc.Release()
	return cookie
}

func (this *OLEControl) UnRegisterEventHandlers(cookie uint32) {
	var cpc *win32.IConnectionPointContainer
	hr := this.QueryInterface(&win32.IID_IConnectionPointContainer, unsafe.Pointer(&cpc))
	win32.ASSERT_SUCCEEDED(hr)

	var cp *win32.IConnectionPoint
	hr = cpc.FindConnectionPoint(&IID_OCXEvents, &cp)
	win32.ASSERT_SUCCEEDED(hr)

	hr = cp.Unadvise(cookie)
	win32.ASSERT_SUCCEEDED(hr)

	cp.Release()
	cpc.Release()
}

