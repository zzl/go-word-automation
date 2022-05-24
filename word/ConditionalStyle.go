package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 1498F56D-ED33-41F9-B37B-EF30E50B08AC
var IID_ConditionalStyle = syscall.GUID{0x1498F56D, 0xED33, 0x41F9, 
	[8]byte{0xB3, 0x7B, 0xEF, 0x30, 0xE5, 0x0B, 0x08, 0xAC}}

type ConditionalStyle struct {
	ole.OleClient
}

func NewConditionalStyle(pDisp *win32.IDispatch, addRef bool, scoped bool) *ConditionalStyle {
	 if pDisp == nil {
		return nil;
	}
	p := &ConditionalStyle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ConditionalStyleFromVar(v ole.Variant) *ConditionalStyle {
	return NewConditionalStyle(v.IDispatch(), false, false)
}

func (this *ConditionalStyle) IID() *syscall.GUID {
	return &IID_ConditionalStyle
}

func (this *ConditionalStyle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ConditionalStyle) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ConditionalStyle) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ConditionalStyle) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000064, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ConditionalStyle) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *ConditionalStyle) Borders() *Borders {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *ConditionalStyle) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *ConditionalStyle) BottomPadding() float32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.FltValVal()
}

func (this *ConditionalStyle) SetBottomPadding(rhs float32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *ConditionalStyle) TopPadding() float32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.FltValVal()
}

func (this *ConditionalStyle) SetTopPadding(rhs float32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *ConditionalStyle) LeftPadding() float32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.FltValVal()
}

func (this *ConditionalStyle) SetLeftPadding(rhs float32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *ConditionalStyle) RightPadding() float32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.FltValVal()
}

func (this *ConditionalStyle) SetRightPadding(rhs float32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *ConditionalStyle) ParagraphFormat() *ParagraphFormat {
	retVal, _ := this.PropGet(0x00000009, nil)
	return NewParagraphFormat(retVal.IDispatch(), false, true)
}

func (this *ConditionalStyle) SetParagraphFormat(rhs *ParagraphFormat)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *ConditionalStyle) Font() *Font {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *ConditionalStyle) SetFont(rhs *Font)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

