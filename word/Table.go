package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020951-0000-0000-C000-000000000046
var IID_Table = syscall.GUID{0x00020951, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Table struct {
	ole.OleClient
}

func NewTable(pDisp *win32.IDispatch, addRef bool, scoped bool) *Table {
	 if pDisp == nil {
		return nil;
	}
	p := &Table{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TableFromVar(v ole.Variant) *Table {
	return NewTable(v.IDispatch(), false, false)
}

func (this *Table) IID() *syscall.GUID {
	return &IID_Table
}

func (this *Table) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Table) Range() *Range {
	retVal, _ := this.PropGet(0x00000000, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Table) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Table) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Table) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Table) Columns() *Columns {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewColumns(retVal.IDispatch(), false, true)
}

func (this *Table) Rows() *Rows {
	retVal, _ := this.PropGet(0x00000065, nil)
	return NewRows(retVal.IDispatch(), false, true)
}

func (this *Table) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Table) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Table) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Table) Uniform() bool {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) AutoFormatType() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *Table) Select()  {
	retVal, _ := this.Call(0x000000c8, nil)
	_= retVal
}

func (this *Table) Delete()  {
	retVal, _ := this.Call(0x00000009, nil)
	_= retVal
}

var Table_SortOld_OptArgs= []string{
	"ExcludeHeader", "FieldNumber", "SortFieldType", "SortOrder", 
	"FieldNumber2", "SortFieldType2", "SortOrder2", "FieldNumber3", 
	"SortFieldType3", "SortOrder3", "CaseSensitive", "LanguageID", 
}

func (this *Table) SortOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Table_SortOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000a, nil, optArgs...)
	_= retVal
}

func (this *Table) SortAscending()  {
	retVal, _ := this.Call(0x0000000c, nil)
	_= retVal
}

func (this *Table) SortDescending()  {
	retVal, _ := this.Call(0x0000000d, nil)
	_= retVal
}

var Table_AutoFormat_OptArgs= []string{
	"Format", "ApplyBorders", "ApplyShading", "ApplyFont", 
	"ApplyColor", "ApplyHeadingRows", "ApplyLastRow", "ApplyFirstColumn", 
	"ApplyLastColumn", "AutoFit", 
}

func (this *Table) AutoFormat(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Table_AutoFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000e, nil, optArgs...)
	_= retVal
}

func (this *Table) UpdateAutoFormat()  {
	retVal, _ := this.Call(0x0000000f, nil)
	_= retVal
}

var Table_ConvertToTextOld_OptArgs= []string{
	"Separator", 
}

func (this *Table) ConvertToTextOld(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Table_ConvertToTextOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000010, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Table) Cell(row int32, column int32) *Cell {
	retVal, _ := this.Call(0x00000011, []interface{}{row, column})
	return NewCell(retVal.IDispatch(), false, true)
}

func (this *Table) Split(beforeRow *ole.Variant) *Table {
	retVal, _ := this.Call(0x00000012, []interface{}{beforeRow})
	return NewTable(retVal.IDispatch(), false, true)
}

var Table_ConvertToText_OptArgs= []string{
	"Separator", "NestedTables", 
}

func (this *Table) ConvertToText(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Table_ConvertToText_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000013, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Table) AutoFitBehavior(behavior int32)  {
	retVal, _ := this.Call(0x00000014, []interface{}{behavior})
	_= retVal
}

var Table_Sort_OptArgs= []string{
	"ExcludeHeader", "FieldNumber", "SortFieldType", "SortOrder", 
	"FieldNumber2", "SortFieldType2", "SortOrder2", "FieldNumber3", 
	"SortFieldType3", "SortOrder3", "CaseSensitive", "BidiSort", 
	"IgnoreThe", "IgnoreKashida", "IgnoreDiacritics", "IgnoreHe", "LanguageID", 
}

func (this *Table) Sort(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Table_Sort_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000017, nil, optArgs...)
	_= retVal
}

func (this *Table) Tables() *Tables {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return NewTables(retVal.IDispatch(), false, true)
}

func (this *Table) NestingLevel() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Table) AllowPageBreaks() bool {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetAllowPageBreaks(rhs bool)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *Table) AllowAutoFit() bool {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetAllowAutoFit(rhs bool)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Table) PreferredWidth() float32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *Table) SetPreferredWidth(rhs float32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *Table) PreferredWidthType() int32 {
	retVal, _ := this.PropGet(0x00000070, nil)
	return retVal.LValVal()
}

func (this *Table) SetPreferredWidthType(rhs int32)  {
	_ = this.PropPut(0x00000070, []interface{}{rhs})
}

func (this *Table) TopPadding() float32 {
	retVal, _ := this.PropGet(0x00000071, nil)
	return retVal.FltValVal()
}

func (this *Table) SetTopPadding(rhs float32)  {
	_ = this.PropPut(0x00000071, []interface{}{rhs})
}

func (this *Table) BottomPadding() float32 {
	retVal, _ := this.PropGet(0x00000072, nil)
	return retVal.FltValVal()
}

func (this *Table) SetBottomPadding(rhs float32)  {
	_ = this.PropPut(0x00000072, []interface{}{rhs})
}

func (this *Table) LeftPadding() float32 {
	retVal, _ := this.PropGet(0x00000073, nil)
	return retVal.FltValVal()
}

func (this *Table) SetLeftPadding(rhs float32)  {
	_ = this.PropPut(0x00000073, []interface{}{rhs})
}

func (this *Table) RightPadding() float32 {
	retVal, _ := this.PropGet(0x00000074, nil)
	return retVal.FltValVal()
}

func (this *Table) SetRightPadding(rhs float32)  {
	_ = this.PropPut(0x00000074, []interface{}{rhs})
}

func (this *Table) Spacing() float32 {
	retVal, _ := this.PropGet(0x00000075, nil)
	return retVal.FltValVal()
}

func (this *Table) SetSpacing(rhs float32)  {
	_ = this.PropPut(0x00000075, []interface{}{rhs})
}

func (this *Table) TableDirection() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Table) SetTableDirection(rhs int32)  {
	_ = this.PropPut(0x00000076, []interface{}{rhs})
}

func (this *Table) ID() string {
	retVal, _ := this.PropGet(0x00000077, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Table) SetID(rhs string)  {
	_ = this.PropPut(0x00000077, []interface{}{rhs})
}

func (this *Table) Style() ole.Variant {
	retVal, _ := this.PropGet(0x000000c9, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Table) SetStyle(rhs *ole.Variant)  {
	_ = this.PropPut(0x000000c9, []interface{}{rhs})
}

func (this *Table) ApplyStyleHeadingRows() bool {
	retVal, _ := this.PropGet(0x000000ca, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleHeadingRows(rhs bool)  {
	_ = this.PropPut(0x000000ca, []interface{}{rhs})
}

func (this *Table) ApplyStyleLastRow() bool {
	retVal, _ := this.PropGet(0x000000cb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleLastRow(rhs bool)  {
	_ = this.PropPut(0x000000cb, []interface{}{rhs})
}

func (this *Table) ApplyStyleFirstColumn() bool {
	retVal, _ := this.PropGet(0x000000cc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleFirstColumn(rhs bool)  {
	_ = this.PropPut(0x000000cc, []interface{}{rhs})
}

func (this *Table) ApplyStyleLastColumn() bool {
	retVal, _ := this.PropGet(0x000000cd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleLastColumn(rhs bool)  {
	_ = this.PropPut(0x000000cd, []interface{}{rhs})
}

func (this *Table) ApplyStyleRowBands() bool {
	retVal, _ := this.PropGet(0x000000ce, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleRowBands(rhs bool)  {
	_ = this.PropPut(0x000000ce, []interface{}{rhs})
}

func (this *Table) ApplyStyleColumnBands() bool {
	retVal, _ := this.PropGet(0x000000cf, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleColumnBands(rhs bool)  {
	_ = this.PropPut(0x000000cf, []interface{}{rhs})
}

func (this *Table) ApplyStyleDirectFormatting(styleName string)  {
	retVal, _ := this.Call(0x000000d0, []interface{}{styleName})
	_= retVal
}

func (this *Table) Title() string {
	retVal, _ := this.PropGet(0x000000d1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Table) SetTitle(rhs string)  {
	_ = this.PropPut(0x000000d1, []interface{}{rhs})
}

func (this *Table) Descr() string {
	retVal, _ := this.PropGet(0x000000d2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Table) SetDescr(rhs string)  {
	_ = this.PropPut(0x000000d2, []interface{}{rhs})
}

