package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 2503B6EE-0889-44DF-B920-6D6F9659DEA3
var IID_OMathBorderBox = syscall.GUID{0x2503B6EE, 0x0889, 0x44DF, 
	[8]byte{0xB9, 0x20, 0x6D, 0x6F, 0x96, 0x59, 0xDE, 0xA3}}

type OMathBorderBox struct {
	ole.OleClient
}

func NewOMathBorderBox(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathBorderBox {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathBorderBox{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathBorderBoxFromVar(v ole.Variant) *OMathBorderBox {
	return NewOMathBorderBox(v.IDispatch(), false, false)
}

func (this *OMathBorderBox) IID() *syscall.GUID {
	return &IID_OMathBorderBox
}

func (this *OMathBorderBox) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathBorderBox) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathBorderBox) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathBorderBox) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathBorderBox) E() *OMath {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathBorderBox) HideTop() bool {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetHideTop(rhs bool)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *OMathBorderBox) HideBot() bool {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetHideBot(rhs bool)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *OMathBorderBox) HideLeft() bool {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetHideLeft(rhs bool)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *OMathBorderBox) HideRight() bool {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetHideRight(rhs bool)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *OMathBorderBox) StrikeH() bool {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetStrikeH(rhs bool)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *OMathBorderBox) StrikeV() bool {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetStrikeV(rhs bool)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *OMathBorderBox) StrikeBLTR() bool {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetStrikeBLTR(rhs bool)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *OMathBorderBox) StrikeTLBR() bool {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetStrikeTLBR(rhs bool)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

