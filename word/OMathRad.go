package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 67A7EEC5-285D-4024-B071-BD6B33B88547
var IID_OMathRad = syscall.GUID{0x67A7EEC5, 0x285D, 0x4024, 
	[8]byte{0xB0, 0x71, 0xBD, 0x6B, 0x33, 0xB8, 0x85, 0x47}}

type OMathRad struct {
	ole.OleClient
}

func NewOMathRad(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathRad {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathRad{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathRadFromVar(v ole.Variant) *OMathRad {
	return NewOMathRad(v.IDispatch(), false, false)
}

func (this *OMathRad) IID() *syscall.GUID {
	return &IID_OMathRad
}

func (this *OMathRad) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathRad) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathRad) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathRad) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathRad) Deg() *OMath {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathRad) E() *OMath {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathRad) HideDeg() bool {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathRad) SetHideDeg(rhs bool)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

