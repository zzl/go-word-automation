package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020950-0000-0000-C000-000000000046
var IID_Row = syscall.GUID{0x00020950, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Row struct {
	ole.OleClient
}

func NewRow(pDisp *win32.IDispatch, addRef bool, scoped bool) *Row {
	 if pDisp == nil {
		return nil;
	}
	p := &Row{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RowFromVar(v ole.Variant) *Row {
	return NewRow(v.IDispatch(), false, false)
}

func (this *Row) IID() *syscall.GUID {
	return &IID_Row
}

func (this *Row) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Row) Range() *Range {
	retVal, _ := this.PropGet(0x00000000, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Row) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Row) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Row) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Row) AllowBreakAcrossPages() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Row) SetAllowBreakAcrossPages(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Row) Alignment() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Row) SetAlignment(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *Row) HeadingFormat() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Row) SetHeadingFormat(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *Row) SpaceBetweenColumns() float32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.FltValVal()
}

func (this *Row) SetSpaceBetweenColumns(rhs float32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *Row) Height() float32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.FltValVal()
}

func (this *Row) SetHeight(rhs float32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *Row) HeightRule() int32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Row) SetHeightRule(rhs int32)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *Row) LeftIndent() float32 {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.FltValVal()
}

func (this *Row) SetLeftIndent(rhs float32)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *Row) IsLast() bool {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Row) IsFirst() bool {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Row) Index() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *Row) Cells() *Cells {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewCells(retVal.IDispatch(), false, true)
}

func (this *Row) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Row) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Row) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Row) Next() *Row {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewRow(retVal.IDispatch(), false, true)
}

func (this *Row) Previous() *Row {
	retVal, _ := this.PropGet(0x00000069, nil)
	return NewRow(retVal.IDispatch(), false, true)
}

func (this *Row) Select()  {
	retVal, _ := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *Row) Delete()  {
	retVal, _ := this.Call(0x000000c8, nil)
	_= retVal
}

func (this *Row) SetLeftIndent_(leftIndent float32, rulerStyle int32)  {
	retVal, _ := this.Call(0x000000ca, []interface{}{leftIndent, rulerStyle})
	_= retVal
}

func (this *Row) SetHeight_(rowHeight float32, heightRule int32)  {
	retVal, _ := this.Call(0x000000cb, []interface{}{rowHeight, heightRule})
	_= retVal
}

var Row_ConvertToTextOld_OptArgs= []string{
	"Separator", 
}

func (this *Row) ConvertToTextOld(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Row_ConvertToTextOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000010, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

var Row_ConvertToText_OptArgs= []string{
	"Separator", "NestedTables", 
}

func (this *Row) ConvertToText(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Row_ConvertToText_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000012, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Row) NestingLevel() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *Row) ID() string {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Row) SetID(rhs string)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

