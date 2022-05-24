package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// CEBD4184-4E6D-4FC6-A42D-2142B1B76AF5
var IID_OMathNary = syscall.GUID{0xCEBD4184, 0x4E6D, 0x4FC6, 
	[8]byte{0xA4, 0x2D, 0x21, 0x42, 0xB1, 0xB7, 0x6A, 0xF5}}

type OMathNary struct {
	ole.OleClient
}

func NewOMathNary(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathNary {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathNary{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathNaryFromVar(v ole.Variant) *OMathNary {
	return NewOMathNary(v.IDispatch(), false, false)
}

func (this *OMathNary) IID() *syscall.GUID {
	return &IID_OMathNary
}

func (this *OMathNary) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathNary) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathNary) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathNary) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathNary) Sub() *OMath {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathNary) Sup() *OMath {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathNary) E() *OMath {
	retVal, _ := this.PropGet(0x00000069, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathNary) Char() int16 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.IValVal()
}

func (this *OMathNary) SetChar(rhs int16)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *OMathNary) Grow() bool {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathNary) SetGrow(rhs bool)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *OMathNary) SubSupLim() bool {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathNary) SetSubSupLim(rhs bool)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *OMathNary) HideSub() bool {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathNary) SetHideSub(rhs bool)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *OMathNary) HideSup() bool {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathNary) SetHideSup(rhs bool)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

