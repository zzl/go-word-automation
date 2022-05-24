package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// DC489AD4-23C4-4F4B-990F-45A51C7C0C4F
var IID_OMathScrSubSup = syscall.GUID{0xDC489AD4, 0x23C4, 0x4F4B, 
	[8]byte{0x99, 0x0F, 0x45, 0xA5, 0x1C, 0x7C, 0x0C, 0x4F}}

type OMathScrSubSup struct {
	ole.OleClient
}

func NewOMathScrSubSup(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathScrSubSup {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathScrSubSup{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathScrSubSupFromVar(v ole.Variant) *OMathScrSubSup {
	return NewOMathScrSubSup(v.IDispatch(), false, false)
}

func (this *OMathScrSubSup) IID() *syscall.GUID {
	return &IID_OMathScrSubSup
}

func (this *OMathScrSubSup) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathScrSubSup) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathScrSubSup) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathScrSubSup) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathScrSubSup) E() *OMath {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathScrSubSup) Sub() *OMath {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathScrSubSup) Sup() *OMath {
	retVal, _ := this.PropGet(0x00000069, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathScrSubSup) AlignScripts() bool {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathScrSubSup) SetAlignScripts(rhs bool)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *OMathScrSubSup) RemoveSub() *OMathFunction {
	retVal, _ := this.Call(0x000000c8, nil)
	return NewOMathFunction(retVal.IDispatch(), false, true)
}

func (this *OMathScrSubSup) RemoveSup() *OMathFunction {
	retVal, _ := this.Call(0x000000c9, nil)
	return NewOMathFunction(retVal.IDispatch(), false, true)
}

func (this *OMathScrSubSup) ToScrPre() *OMathFunction {
	retVal, _ := this.Call(0x000000ca, nil)
	return NewOMathFunction(retVal.IDispatch(), false, true)
}

