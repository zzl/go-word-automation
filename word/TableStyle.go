package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// B7564E97-0519-4C68-B400-3803E7C63242
var IID_TableStyle = syscall.GUID{0xB7564E97, 0x0519, 0x4C68, 
	[8]byte{0xB4, 0x00, 0x38, 0x03, 0xE7, 0xC6, 0x32, 0x42}}

type TableStyle struct {
	ole.OleClient
}

func NewTableStyle(pDisp *win32.IDispatch, addRef bool, scoped bool) *TableStyle {
	 if pDisp == nil {
		return nil;
	}
	p := &TableStyle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TableStyleFromVar(v ole.Variant) *TableStyle {
	return NewTableStyle(v.IDispatch(), false, false)
}

func (this *TableStyle) IID() *syscall.GUID {
	return &IID_TableStyle
}

func (this *TableStyle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TableStyle) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TableStyle) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TableStyle) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000064, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TableStyle) AllowPageBreaks() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableStyle) SetAllowPageBreaks(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *TableStyle) Borders() *Borders {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *TableStyle) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *TableStyle) BottomPadding() float32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.FltValVal()
}

func (this *TableStyle) SetBottomPadding(rhs float32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *TableStyle) LeftPadding() float32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.FltValVal()
}

func (this *TableStyle) SetLeftPadding(rhs float32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *TableStyle) TopPadding() float32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.FltValVal()
}

func (this *TableStyle) SetTopPadding(rhs float32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *TableStyle) RightPadding() float32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.FltValVal()
}

func (this *TableStyle) SetRightPadding(rhs float32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *TableStyle) Alignment() int32 {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.LValVal()
}

func (this *TableStyle) SetAlignment(rhs int32)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *TableStyle) Spacing() float32 {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.FltValVal()
}

func (this *TableStyle) SetSpacing(rhs float32)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *TableStyle) Condition(conditionCode int32) *ConditionalStyle {
	retVal, _ := this.Call(0x00000010, []interface{}{conditionCode})
	return NewConditionalStyle(retVal.IDispatch(), false, true)
}

func (this *TableStyle) TableDirection() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *TableStyle) SetTableDirection(rhs int32)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *TableStyle) AllowBreakAcrossPage() int32 {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.LValVal()
}

func (this *TableStyle) SetAllowBreakAcrossPage(rhs int32)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *TableStyle) LeftIndent() float32 {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return retVal.FltValVal()
}

func (this *TableStyle) SetLeftIndent(rhs float32)  {
	_ = this.PropPut(0x0000000e, []interface{}{rhs})
}

func (this *TableStyle) Shading() *Shading {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *TableStyle) RowStripe() int32 {
	retVal, _ := this.PropGet(0x00000011, nil)
	return retVal.LValVal()
}

func (this *TableStyle) SetRowStripe(rhs int32)  {
	_ = this.PropPut(0x00000011, []interface{}{rhs})
}

func (this *TableStyle) ColumnStripe() int32 {
	retVal, _ := this.PropGet(0x00000012, nil)
	return retVal.LValVal()
}

func (this *TableStyle) SetColumnStripe(rhs int32)  {
	_ = this.PropPut(0x00000012, []interface{}{rhs})
}

