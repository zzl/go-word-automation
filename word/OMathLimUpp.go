package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// FC9086C6-0287-4997-B2E1-816C334A22F8
var IID_OMathLimUpp = syscall.GUID{0xFC9086C6, 0x0287, 0x4997, 
	[8]byte{0xB2, 0xE1, 0x81, 0x6C, 0x33, 0x4A, 0x22, 0xF8}}

type OMathLimUpp struct {
	ole.OleClient
}

func NewOMathLimUpp(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathLimUpp {
	p := &OMathLimUpp{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathLimUppFromVar(v ole.Variant) *OMathLimUpp {
	return NewOMathLimUpp(v.PdispValVal(), false, false)
}

func (this *OMathLimUpp) IID() *syscall.GUID {
	return &IID_OMathLimUpp
}

func (this *OMathLimUpp) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathLimUpp) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathLimUpp) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathLimUpp) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathLimUpp) E() *OMath {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathLimUpp) Lim() *OMath {
	retVal := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathLimUpp) ToLimLow() *OMathFunction {
	retVal := this.Call(0x000000c9, nil)
	return NewOMathFunction(retVal.PdispValVal(), false, true)
}

