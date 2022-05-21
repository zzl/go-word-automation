package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 74DE9576-8E99-4E28-912B-CB30747C60CE
var IID_OMathLimLow = syscall.GUID{0x74DE9576, 0x8E99, 0x4E28, 
	[8]byte{0x91, 0x2B, 0xCB, 0x30, 0x74, 0x7C, 0x60, 0xCE}}

type OMathLimLow struct {
	ole.OleClient
}

func NewOMathLimLow(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathLimLow {
	p := &OMathLimLow{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathLimLowFromVar(v ole.Variant) *OMathLimLow {
	return NewOMathLimLow(v.PdispValVal(), false, false)
}

func (this *OMathLimLow) IID() *syscall.GUID {
	return &IID_OMathLimLow
}

func (this *OMathLimLow) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathLimLow) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathLimLow) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathLimLow) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathLimLow) E() *OMath {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathLimLow) Lim() *OMath {
	retVal := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathLimLow) ToLimUpp() *OMathFunction {
	retVal := this.Call(0x000000c9, nil)
	return NewOMathFunction(retVal.PdispValVal(), false, true)
}

