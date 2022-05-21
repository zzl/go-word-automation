package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209F3-0000-0000-C000-000000000046
var IID_OCXEvents = syscall.GUID{0x000209F3, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OCXEventsDispInterface interface {
	GotFocus() 
	LostFocus() 
}

type OCXEventsHandlers struct {
	GotFocus func() 
	LostFocus func() 
}

type OCXEventsDispImpl struct {
	Handlers OCXEventsHandlers
}

func (this *OCXEventsDispImpl) GotFocus() {
	if this.Handlers.GotFocus != nil {
		this.Handlers.GotFocus()
	}
}

func (this *OCXEventsDispImpl) LostFocus() {
	if this.Handlers.LostFocus != nil {
		this.Handlers.LostFocus()
	}
}

type OCXEventsImpl struct {
	ole.IDispatchImpl
	DispImpl OCXEventsDispInterface
}

func (this *OCXEventsImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == IID_OCXEvents {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IDispatchImpl.QueryInterface(riid, ppvObject)
}

func (this *OCXEventsImpl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32,
	wFlags uint16, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT,
	pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {
	var unwrapActions ole.Actions
	defer unwrapActions.Execute()
	switch dispIdMember {
	case -2147417888:
		this.DispImpl.GotFocus()
		return win32.S_OK
	case -2147417887:
		this.DispImpl.LostFocus()
		return win32.S_OK
	}
	return win32.E_NOTIMPL
}

type OCXEventsComObj struct {
	ole.IDispatchComObj
}

func NewOCXEventsComObj(dispImpl OCXEventsDispInterface, scoped bool) *OCXEventsComObj {
	comObj := com.NewComObj[OCXEventsComObj](
		&OCXEventsImpl {DispImpl: dispImpl})
	if scoped {
		com.AddToScope(comObj)
	}
	return comObj
}

