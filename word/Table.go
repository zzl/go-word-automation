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
	return NewTable(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x00000000, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Table) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Table) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Table) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Table) Columns() *Columns {
	retVal := this.PropGet(0x00000064, nil)
	return NewColumns(retVal.PdispValVal(), false, true)
}

func (this *Table) Rows() *Rows {
	retVal := this.PropGet(0x00000065, nil)
	return NewRows(retVal.PdispValVal(), false, true)
}

func (this *Table) Borders() *Borders {
	retVal := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *Table) SetBorders(rhs *Borders)  {
	retVal := this.PropPut(0x0000044c, []interface{}{rhs})
	_= retVal
}

func (this *Table) Shading() *Shading {
	retVal := this.PropGet(0x00000068, nil)
	return NewShading(retVal.PdispValVal(), false, true)
}

func (this *Table) Uniform() bool {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) AutoFormatType() int32 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *Table) Select()  {
	retVal := this.Call(0x000000c8, nil)
	_= retVal
}

func (this *Table) Delete()  {
	retVal := this.Call(0x00000009, nil)
	_= retVal
}

var Table_SortOld_OptArgs= []string{
	"ExcludeHeader", "FieldNumber", "SortFieldType", "SortOrder", 
	"FieldNumber2", "SortFieldType2", "SortOrder2", "FieldNumber3", 
	"SortFieldType3", "SortOrder3", "CaseSensitive", "LanguageID", 
}

func (this *Table) SortOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Table_SortOld_OptArgs, optArgs)
	retVal := this.Call(0x0000000a, nil, optArgs...)
	_= retVal
}

func (this *Table) SortAscending()  {
	retVal := this.Call(0x0000000c, nil)
	_= retVal
}

func (this *Table) SortDescending()  {
	retVal := this.Call(0x0000000d, nil)
	_= retVal
}

var Table_AutoFormat_OptArgs= []string{
	"Format", "ApplyBorders", "ApplyShading", "ApplyFont", 
	"ApplyColor", "ApplyHeadingRows", "ApplyLastRow", "ApplyFirstColumn", 
	"ApplyLastColumn", "AutoFit", 
}

func (this *Table) AutoFormat(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Table_AutoFormat_OptArgs, optArgs)
	retVal := this.Call(0x0000000e, nil, optArgs...)
	_= retVal
}

func (this *Table) UpdateAutoFormat()  {
	retVal := this.Call(0x0000000f, nil)
	_= retVal
}

var Table_ConvertToTextOld_OptArgs= []string{
	"Separator", 
}

func (this *Table) ConvertToTextOld(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Table_ConvertToTextOld_OptArgs, optArgs)
	retVal := this.Call(0x00000010, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Table) Cell(row int32, column int32) *Cell {
	retVal := this.Call(0x00000011, []interface{}{row, column})
	return NewCell(retVal.PdispValVal(), false, true)
}

func (this *Table) Split(beforeRow *ole.Variant) *Table {
	retVal := this.Call(0x00000012, []interface{}{beforeRow})
	return NewTable(retVal.PdispValVal(), false, true)
}

var Table_ConvertToText_OptArgs= []string{
	"Separator", "NestedTables", 
}

func (this *Table) ConvertToText(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Table_ConvertToText_OptArgs, optArgs)
	retVal := this.Call(0x00000013, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Table) AutoFitBehavior(behavior int32)  {
	retVal := this.Call(0x00000014, []interface{}{behavior})
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
	retVal := this.Call(0x00000017, nil, optArgs...)
	_= retVal
}

func (this *Table) Tables() *Tables {
	retVal := this.PropGet(0x0000006b, nil)
	return NewTables(retVal.PdispValVal(), false, true)
}

func (this *Table) NestingLevel() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Table) AllowPageBreaks() bool {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetAllowPageBreaks(rhs bool)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *Table) AllowAutoFit() bool {
	retVal := this.PropGet(0x0000006e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetAllowAutoFit(rhs bool)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *Table) PreferredWidth() float32 {
	retVal := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *Table) SetPreferredWidth(rhs float32)  {
	retVal := this.PropPut(0x0000006f, []interface{}{rhs})
	_= retVal
}

func (this *Table) PreferredWidthType() int32 {
	retVal := this.PropGet(0x00000070, nil)
	return retVal.LValVal()
}

func (this *Table) SetPreferredWidthType(rhs int32)  {
	retVal := this.PropPut(0x00000070, []interface{}{rhs})
	_= retVal
}

func (this *Table) TopPadding() float32 {
	retVal := this.PropGet(0x00000071, nil)
	return retVal.FltValVal()
}

func (this *Table) SetTopPadding(rhs float32)  {
	retVal := this.PropPut(0x00000071, []interface{}{rhs})
	_= retVal
}

func (this *Table) BottomPadding() float32 {
	retVal := this.PropGet(0x00000072, nil)
	return retVal.FltValVal()
}

func (this *Table) SetBottomPadding(rhs float32)  {
	retVal := this.PropPut(0x00000072, []interface{}{rhs})
	_= retVal
}

func (this *Table) LeftPadding() float32 {
	retVal := this.PropGet(0x00000073, nil)
	return retVal.FltValVal()
}

func (this *Table) SetLeftPadding(rhs float32)  {
	retVal := this.PropPut(0x00000073, []interface{}{rhs})
	_= retVal
}

func (this *Table) RightPadding() float32 {
	retVal := this.PropGet(0x00000074, nil)
	return retVal.FltValVal()
}

func (this *Table) SetRightPadding(rhs float32)  {
	retVal := this.PropPut(0x00000074, []interface{}{rhs})
	_= retVal
}

func (this *Table) Spacing() float32 {
	retVal := this.PropGet(0x00000075, nil)
	return retVal.FltValVal()
}

func (this *Table) SetSpacing(rhs float32)  {
	retVal := this.PropPut(0x00000075, []interface{}{rhs})
	_= retVal
}

func (this *Table) TableDirection() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Table) SetTableDirection(rhs int32)  {
	retVal := this.PropPut(0x00000076, []interface{}{rhs})
	_= retVal
}

func (this *Table) ID() string {
	retVal := this.PropGet(0x00000077, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Table) SetID(rhs string)  {
	retVal := this.PropPut(0x00000077, []interface{}{rhs})
	_= retVal
}

func (this *Table) Style() ole.Variant {
	retVal := this.PropGet(0x000000c9, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Table) SetStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x000000c9, []interface{}{rhs})
	_= retVal
}

func (this *Table) ApplyStyleHeadingRows() bool {
	retVal := this.PropGet(0x000000ca, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleHeadingRows(rhs bool)  {
	retVal := this.PropPut(0x000000ca, []interface{}{rhs})
	_= retVal
}

func (this *Table) ApplyStyleLastRow() bool {
	retVal := this.PropGet(0x000000cb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleLastRow(rhs bool)  {
	retVal := this.PropPut(0x000000cb, []interface{}{rhs})
	_= retVal
}

func (this *Table) ApplyStyleFirstColumn() bool {
	retVal := this.PropGet(0x000000cc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleFirstColumn(rhs bool)  {
	retVal := this.PropPut(0x000000cc, []interface{}{rhs})
	_= retVal
}

func (this *Table) ApplyStyleLastColumn() bool {
	retVal := this.PropGet(0x000000cd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleLastColumn(rhs bool)  {
	retVal := this.PropPut(0x000000cd, []interface{}{rhs})
	_= retVal
}

func (this *Table) ApplyStyleRowBands() bool {
	retVal := this.PropGet(0x000000ce, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleRowBands(rhs bool)  {
	retVal := this.PropPut(0x000000ce, []interface{}{rhs})
	_= retVal
}

func (this *Table) ApplyStyleColumnBands() bool {
	retVal := this.PropGet(0x000000cf, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Table) SetApplyStyleColumnBands(rhs bool)  {
	retVal := this.PropPut(0x000000cf, []interface{}{rhs})
	_= retVal
}

func (this *Table) ApplyStyleDirectFormatting(styleName string)  {
	retVal := this.Call(0x000000d0, []interface{}{styleName})
	_= retVal
}

func (this *Table) Title() string {
	retVal := this.PropGet(0x000000d1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Table) SetTitle(rhs string)  {
	retVal := this.PropPut(0x000000d1, []interface{}{rhs})
	_= retVal
}

func (this *Table) Descr() string {
	retVal := this.PropGet(0x000000d2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Table) SetDescr(rhs string)  {
	retVal := this.PropPut(0x000000d2, []interface{}{rhs})
	_= retVal
}

