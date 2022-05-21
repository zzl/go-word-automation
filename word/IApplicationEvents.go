package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209F7-0001-0000-C000-000000000046
var IID_IApplicationEvents = syscall.GUID{0x000209F7, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IApplicationEvents struct {
	ole.OleClient
}

func NewIApplicationEvents(pDisp *win32.IDispatch, addRef bool, scoped bool) *IApplicationEvents {
	p := &IApplicationEvents{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func IApplicationEventsFromVar(v ole.Variant) *IApplicationEvents {
	return NewIApplicationEvents(v.PdispValVal(), false, false)
}

func (this *IApplicationEvents) IID() *syscall.GUID {
	return &IID_IApplicationEvents
}

func (this *IApplicationEvents) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *IApplicationEvents) Startup()  {
	retVal := this.Call(0x00000001, nil)
	_= retVal
}

func (this *IApplicationEvents) Quit()  {
	retVal := this.Call(0x00000002, nil)
	_= retVal
}

func (this *IApplicationEvents) DocumentChange()  {
	retVal := this.Call(0x00000003, nil)
	_= retVal
}

