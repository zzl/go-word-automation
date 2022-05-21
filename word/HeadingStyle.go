package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002098B-0000-0000-C000-000000000046
var IID_HeadingStyle = syscall.GUID{0x0002098B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HeadingStyle struct {
	ole.OleClient
}

func NewHeadingStyle(pDisp *win32.IDispatch, addRef bool, scoped bool) *HeadingStyle {
	p := &HeadingStyle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HeadingStyleFromVar(v ole.Variant) *HeadingStyle {
	return NewHeadingStyle(v.PdispValVal(), false, false)
}

func (this *HeadingStyle) IID() *syscall.GUID {
	return &IID_HeadingStyle
}

func (this *HeadingStyle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HeadingStyle) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *HeadingStyle) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HeadingStyle) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *HeadingStyle) Style() ole.Variant {
	retVal := this.PropGet(0x00000000, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *HeadingStyle) SetStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *HeadingStyle) Level() int16 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.IValVal()
}

func (this *HeadingStyle) SetLevel(rhs int16)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *HeadingStyle) Delete()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

