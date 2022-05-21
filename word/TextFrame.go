package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209B2-0000-0000-C000-000000000046
var IID_TextFrame = syscall.GUID{0x000209B2, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TextFrame struct {
	ole.OleClient
}

func NewTextFrame(pDisp *win32.IDispatch, addRef bool, scoped bool) *TextFrame {
	p := &TextFrame{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TextFrameFromVar(v ole.Variant) *TextFrame {
	return NewTextFrame(v.PdispValVal(), false, false)
}

func (this *TextFrame) IID() *syscall.GUID {
	return &IID_TextFrame
}

func (this *TextFrame) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TextFrame) Application() *Application {
	retVal := this.PropGet(0x00001f40, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *TextFrame) Creator() int32 {
	retVal := this.PropGet(0x00001f41, nil)
	return retVal.LValVal()
}

func (this *TextFrame) Parent() *Shape {
	retVal := this.PropGet(0x00000001, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *TextFrame) MarginBottom() float32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.FltValVal()
}

func (this *TextFrame) SetMarginBottom(rhs float32)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) MarginLeft() float32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.FltValVal()
}

func (this *TextFrame) SetMarginLeft(rhs float32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) MarginRight() float32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *TextFrame) SetMarginRight(rhs float32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) MarginTop() float32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *TextFrame) SetMarginTop(rhs float32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) Orientation() int32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetOrientation(rhs int32)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) TextRange() *Range {
	retVal := this.PropGet(0x000003e9, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *TextFrame) ContainingRange() *Range {
	retVal := this.PropGet(0x000003ea, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *TextFrame) Next() *TextFrame {
	retVal := this.PropGet(0x00001389, nil)
	return NewTextFrame(retVal.PdispValVal(), false, true)
}

func (this *TextFrame) SetNext(rhs *TextFrame)  {
	retVal := this.PropPut(0x00001389, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) Previous() *TextFrame {
	retVal := this.PropGet(0x0000138a, nil)
	return NewTextFrame(retVal.PdispValVal(), false, true)
}

func (this *TextFrame) SetPrevious(rhs *TextFrame)  {
	retVal := this.PropPut(0x0000138a, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) Overflowing() bool {
	retVal := this.PropGet(0x0000138b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextFrame) HasText() int32 {
	retVal := this.PropGet(0x00001390, nil)
	return retVal.LValVal()
}

func (this *TextFrame) BreakForwardLink()  {
	retVal := this.Call(0x0000138c, nil)
	_= retVal
}

func (this *TextFrame) ValidLinkTarget(targetTextFrame *TextFrame) bool {
	retVal := this.Call(0x0000138e, []interface{}{targetTextFrame})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextFrame) AutoSize() int32 {
	retVal := this.PropGet(0x00001391, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetAutoSize(rhs int32)  {
	retVal := this.PropPut(0x00001391, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) WordWrap() int32 {
	retVal := this.PropGet(0x00001392, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetWordWrap(rhs int32)  {
	retVal := this.PropPut(0x00001392, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) VerticalAnchor() int32 {
	retVal := this.PropGet(0x00001393, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetVerticalAnchor(rhs int32)  {
	retVal := this.PropPut(0x00001393, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) HorizontalAnchor() int32 {
	retVal := this.PropGet(0x00001394, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetHorizontalAnchor(rhs int32)  {
	retVal := this.PropPut(0x00001394, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) PathFormat() int32 {
	retVal := this.PropGet(0x00001395, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetPathFormat(rhs int32)  {
	retVal := this.PropPut(0x00001395, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) WarpFormat() int32 {
	retVal := this.PropGet(0x00001396, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetWarpFormat(rhs int32)  {
	retVal := this.PropPut(0x00001396, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) Column() *ole.DispatchClass {
	retVal := this.PropGet(0x00001397, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *TextFrame) ThreeD() *ThreeDFormat {
	retVal := this.PropGet(0x00001398, nil)
	return NewThreeDFormat(retVal.PdispValVal(), false, true)
}

func (this *TextFrame) NoTextRotation() int32 {
	retVal := this.PropGet(0x00001399, nil)
	return retVal.LValVal()
}

func (this *TextFrame) SetNoTextRotation(rhs int32)  {
	retVal := this.PropPut(0x00001399, []interface{}{rhs})
	_= retVal
}

func (this *TextFrame) DeleteText()  {
	retVal := this.Call(0x0000139a, nil)
	_= retVal
}

