package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002092E-0000-0000-C000-000000000046
var IID_Browser = syscall.GUID{0x0002092E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Browser struct {
	ole.OleClient
}

func NewBrowser(pDisp *win32.IDispatch, addRef bool, scoped bool) *Browser {
	p := &Browser{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BrowserFromVar(v ole.Variant) *Browser {
	return NewBrowser(v.PdispValVal(), false, false)
}

func (this *Browser) IID() *syscall.GUID {
	return &IID_Browser
}

func (this *Browser) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Browser) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Browser) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Browser) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Browser) Target() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Browser) SetTarget(rhs int32)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *Browser) Next()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *Browser) Previous()  {
	retVal := this.Call(0x00000066, nil)
	_= retVal
}

