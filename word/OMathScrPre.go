package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// AFAF0C0E-8603-40F6-8FD1-42726CAC21E3
var IID_OMathScrPre = syscall.GUID{0xAFAF0C0E, 0x8603, 0x40F6, 
	[8]byte{0x8F, 0xD1, 0x42, 0x72, 0x6C, 0xAC, 0x21, 0xE3}}

type OMathScrPre struct {
	ole.OleClient
}

func NewOMathScrPre(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathScrPre {
	p := &OMathScrPre{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathScrPreFromVar(v ole.Variant) *OMathScrPre {
	return NewOMathScrPre(v.PdispValVal(), false, false)
}

func (this *OMathScrPre) IID() *syscall.GUID {
	return &IID_OMathScrPre
}

func (this *OMathScrPre) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathScrPre) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathScrPre) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathScrPre) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathScrPre) Sub() *OMath {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathScrPre) Sup() *OMath {
	retVal := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathScrPre) E() *OMath {
	retVal := this.PropGet(0x00000069, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathScrPre) ToScrSubSup() *OMathFunction {
	retVal := this.Call(0x000000c9, nil)
	return NewOMathFunction(retVal.PdispValVal(), false, true)
}

