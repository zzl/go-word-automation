package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// E4442A83-F623-459C-8E95-8BFB44DCF23A
var IID_OMath = syscall.GUID{0xE4442A83, 0xF623, 0x459C, 
	[8]byte{0x8E, 0x95, 0x8B, 0xFB, 0x44, 0xDC, 0xF2, 0x3A}}

type OMath struct {
	ole.OleClient
}

func NewOMath(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMath {
	p := &OMath{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathFromVar(v ole.Variant) *OMath {
	return NewOMath(v.PdispValVal(), false, false)
}

func (this *OMath) IID() *syscall.GUID {
	return &IID_OMath
}

func (this *OMath) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMath) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMath) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMath) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMath) Range() *Range {
	retVal := this.PropGet(0x00000067, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *OMath) Functions() *OMathFunctions {
	retVal := this.PropGet(0x00000068, nil)
	return NewOMathFunctions(retVal.PdispValVal(), false, true)
}

func (this *OMath) Type() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *OMath) SetType(rhs int32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *OMath) ParentOMath() *OMath {
	retVal := this.PropGet(0x0000006a, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMath) ParentFunction() *OMathFunction {
	retVal := this.PropGet(0x0000006b, nil)
	return NewOMathFunction(retVal.PdispValVal(), false, true)
}

func (this *OMath) ParentRow() *OMathMatRow {
	retVal := this.PropGet(0x0000006c, nil)
	return NewOMathMatRow(retVal.PdispValVal(), false, true)
}

func (this *OMath) ParentCol() *OMathMatCol {
	retVal := this.PropGet(0x0000006d, nil)
	return NewOMathMatCol(retVal.PdispValVal(), false, true)
}

func (this *OMath) ParentArg() *OMath {
	retVal := this.PropGet(0x0000006e, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMath) ArgIndex() int32 {
	retVal := this.PropGet(0x0000006f, nil)
	return retVal.LValVal()
}

func (this *OMath) NestingLevel() int32 {
	retVal := this.PropGet(0x00000070, nil)
	return retVal.LValVal()
}

func (this *OMath) ArgSize() int32 {
	retVal := this.PropGet(0x00000071, nil)
	return retVal.LValVal()
}

func (this *OMath) SetArgSize(rhs int32)  {
	retVal := this.PropPut(0x00000071, []interface{}{rhs})
	_= retVal
}

func (this *OMath) Breaks() *OMathBreaks {
	retVal := this.PropGet(0x00000072, nil)
	return NewOMathBreaks(retVal.PdispValVal(), false, true)
}

func (this *OMath) Justification() int32 {
	retVal := this.PropGet(0x00000073, nil)
	return retVal.LValVal()
}

func (this *OMath) SetJustification(rhs int32)  {
	retVal := this.PropPut(0x00000073, []interface{}{rhs})
	_= retVal
}

func (this *OMath) AlignPoint() int32 {
	retVal := this.PropGet(0x00000074, nil)
	return retVal.LValVal()
}

func (this *OMath) SetAlignPoint(rhs int32)  {
	retVal := this.PropPut(0x00000074, []interface{}{rhs})
	_= retVal
}

func (this *OMath) Linearize()  {
	retVal := this.Call(0x000000c8, nil)
	_= retVal
}

func (this *OMath) BuildUp()  {
	retVal := this.Call(0x000000c9, nil)
	_= retVal
}

func (this *OMath) Remove()  {
	retVal := this.Call(0x000000ca, nil)
	_= retVal
}

func (this *OMath) ConvertToMathText()  {
	retVal := this.Call(0x000000cb, nil)
	_= retVal
}

func (this *OMath) ConvertToNormalText()  {
	retVal := this.Call(0x000000cc, nil)
	_= retVal
}

func (this *OMath) ConvertToLiteralText()  {
	retVal := this.Call(0x000000cd, nil)
	_= retVal
}

