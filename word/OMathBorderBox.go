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
	return NewOMathBorderBox(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathBorderBox) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathBorderBox) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathBorderBox) E() *OMath {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathBorderBox) HideTop() bool {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetHideTop(rhs bool)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *OMathBorderBox) HideBot() bool {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetHideBot(rhs bool)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *OMathBorderBox) HideLeft() bool {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetHideLeft(rhs bool)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *OMathBorderBox) HideRight() bool {
	retVal := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetHideRight(rhs bool)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

func (this *OMathBorderBox) StrikeH() bool {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetStrikeH(rhs bool)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *OMathBorderBox) StrikeV() bool {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetStrikeV(rhs bool)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *OMathBorderBox) StrikeBLTR() bool {
	retVal := this.PropGet(0x0000006e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetStrikeBLTR(rhs bool)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *OMathBorderBox) StrikeTLBR() bool {
	retVal := this.PropGet(0x0000006f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBorderBox) SetStrikeTLBR(rhs bool)  {
	retVal := this.PropPut(0x0000006f, []interface{}{rhs})
	_= retVal
}

