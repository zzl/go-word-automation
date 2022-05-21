package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 98DFBD12-96CB-4F07-90EA-749FF1D6B89D
var IID_OMathScrSub = syscall.GUID{0x98DFBD12, 0x96CB, 0x4F07, 
	[8]byte{0x90, 0xEA, 0x74, 0x9F, 0xF1, 0xD6, 0xB8, 0x9D}}

type OMathScrSub struct {
	ole.OleClient
}

func NewOMathScrSub(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathScrSub {
	p := &OMathScrSub{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathScrSubFromVar(v ole.Variant) *OMathScrSub {
	return NewOMathScrSub(v.PdispValVal(), false, false)
}

func (this *OMathScrSub) IID() *syscall.GUID {
	return &IID_OMathScrSub
}

func (this *OMathScrSub) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathScrSub) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathScrSub) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathScrSub) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathScrSub) E() *OMath {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathScrSub) Sub() *OMath {
	retVal := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

