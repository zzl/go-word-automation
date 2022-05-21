package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002092A-0000-0000-C000-000000000046
var IID_Frame = syscall.GUID{0x0002092A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Frame struct {
	ole.OleClient
}

func NewFrame(pDisp *win32.IDispatch, addRef bool, scoped bool) *Frame {
	p := &Frame{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FrameFromVar(v ole.Variant) *Frame {
	return NewFrame(v.PdispValVal(), false, false)
}

func (this *Frame) IID() *syscall.GUID {
	return &IID_Frame
}

func (this *Frame) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Frame) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Frame) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Frame) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Frame) HeightRule() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Frame) SetHeightRule(rhs int32)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *Frame) WidthRule() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Frame) SetWidthRule(rhs int32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *Frame) HorizontalDistanceFromText() float32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.FltValVal()
}

func (this *Frame) SetHorizontalDistanceFromText(rhs float32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *Frame) Height() float32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.FltValVal()
}

func (this *Frame) SetHeight(rhs float32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *Frame) HorizontalPosition() float32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.FltValVal()
}

func (this *Frame) SetHorizontalPosition(rhs float32)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *Frame) LockAnchor() bool {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Frame) SetLockAnchor(rhs bool)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *Frame) RelativeHorizontalPosition() int32 {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *Frame) SetRelativeHorizontalPosition(rhs int32)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *Frame) RelativeVerticalPosition() int32 {
	retVal := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Frame) SetRelativeVerticalPosition(rhs int32)  {
	retVal := this.PropPut(0x00000008, []interface{}{rhs})
	_= retVal
}

func (this *Frame) VerticalDistanceFromText() float32 {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.FltValVal()
}

func (this *Frame) SetVerticalDistanceFromText(rhs float32)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *Frame) VerticalPosition() float32 {
	retVal := this.PropGet(0x0000000a, nil)
	return retVal.FltValVal()
}

func (this *Frame) SetVerticalPosition(rhs float32)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *Frame) Width() float32 {
	retVal := this.PropGet(0x0000000b, nil)
	return retVal.FltValVal()
}

func (this *Frame) SetWidth(rhs float32)  {
	retVal := this.PropPut(0x0000000b, []interface{}{rhs})
	_= retVal
}

func (this *Frame) TextWrap() bool {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Frame) SetTextWrap(rhs bool)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *Frame) Shading() *Shading {
	retVal := this.PropGet(0x0000000d, nil)
	return NewShading(retVal.PdispValVal(), false, true)
}

func (this *Frame) Borders() *Borders {
	retVal := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *Frame) SetBorders(rhs *Borders)  {
	retVal := this.PropPut(0x0000044c, []interface{}{rhs})
	_= retVal
}

func (this *Frame) Range() *Range {
	retVal := this.PropGet(0x0000000f, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Frame) Delete()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

func (this *Frame) Select()  {
	retVal := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *Frame) Copy()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *Frame) Cut()  {
	retVal := this.Call(0x00000066, nil)
	_= retVal
}

