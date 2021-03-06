package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002094E-0000-0000-C000-000000000046
var IID_Cell = syscall.GUID{0x0002094E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Cell struct {
	ole.OleClient
}

func NewCell(pDisp *win32.IDispatch, addRef bool, scoped bool) *Cell {
	 if pDisp == nil {
		return nil;
	}
	p := &Cell{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CellFromVar(v ole.Variant) *Cell {
	return NewCell(v.IDispatch(), false, false)
}

func (this *Cell) IID() *syscall.GUID {
	return &IID_Cell
}

func (this *Cell) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Cell) Range() *Range {
	retVal, _ := this.PropGet(0x00000000, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Cell) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Cell) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Cell) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Cell) RowIndex() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Cell) ColumnIndex() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Cell) Width() float32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.FltValVal()
}

func (this *Cell) SetWidth(rhs float32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *Cell) Height() float32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.FltValVal()
}

func (this *Cell) SetHeight(rhs float32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *Cell) HeightRule() int32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Cell) SetHeightRule(rhs int32)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *Cell) VerticalAlignment() int32 {
	retVal, _ := this.PropGet(0x00000450, nil)
	return retVal.LValVal()
}

func (this *Cell) SetVerticalAlignment(rhs int32)  {
	_ = this.PropPut(0x00000450, []interface{}{rhs})
}

func (this *Cell) Column() *Column {
	retVal, _ := this.PropGet(0x00000065, nil)
	return NewColumn(retVal.IDispatch(), false, true)
}

func (this *Cell) Row() *Row {
	retVal, _ := this.PropGet(0x00000066, nil)
	return NewRow(retVal.IDispatch(), false, true)
}

func (this *Cell) Next() *Cell {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewCell(retVal.IDispatch(), false, true)
}

func (this *Cell) Previous() *Cell {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewCell(retVal.IDispatch(), false, true)
}

func (this *Cell) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000069, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Cell) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Cell) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Cell) Select()  {
	retVal, _ := this.Call(0x0000ffff, nil)
	_= retVal
}

var Cell_Delete_OptArgs= []string{
	"ShiftCells", 
}

func (this *Cell) Delete(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Cell_Delete_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000c8, nil, optArgs...)
	_= retVal
}

var Cell_Formula_OptArgs= []string{
	"Formula", "NumFormat", 
}

func (this *Cell) Formula(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Cell_Formula_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000c9, nil, optArgs...)
	_= retVal
}

func (this *Cell) SetWidth_(columnWidth float32, rulerStyle int32)  {
	retVal, _ := this.Call(0x000000ca, []interface{}{columnWidth, rulerStyle})
	_= retVal
}

func (this *Cell) SetHeight_(rowHeight *ole.Variant, heightRule int32)  {
	retVal, _ := this.Call(0x000000cb, []interface{}{rowHeight, heightRule})
	_= retVal
}

func (this *Cell) Merge(mergeTo *Cell)  {
	retVal, _ := this.Call(0x000000cc, []interface{}{mergeTo})
	_= retVal
}

var Cell_Split_OptArgs= []string{
	"NumRows", "NumColumns", 
}

func (this *Cell) Split(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Cell_Split_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000cd, nil, optArgs...)
	_= retVal
}

func (this *Cell) AutoSum()  {
	retVal, _ := this.Call(0x000000ce, nil)
	_= retVal
}

func (this *Cell) Tables() *Tables {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return NewTables(retVal.IDispatch(), false, true)
}

func (this *Cell) NestingLevel() int32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *Cell) WordWrap() bool {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Cell) SetWordWrap(rhs bool)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *Cell) PreferredWidth() float32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.FltValVal()
}

func (this *Cell) SetPreferredWidth(rhs float32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *Cell) FitText() bool {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Cell) SetFitText(rhs bool)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Cell) TopPadding() float32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *Cell) SetTopPadding(rhs float32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *Cell) BottomPadding() float32 {
	retVal, _ := this.PropGet(0x00000070, nil)
	return retVal.FltValVal()
}

func (this *Cell) SetBottomPadding(rhs float32)  {
	_ = this.PropPut(0x00000070, []interface{}{rhs})
}

func (this *Cell) LeftPadding() float32 {
	retVal, _ := this.PropGet(0x00000071, nil)
	return retVal.FltValVal()
}

func (this *Cell) SetLeftPadding(rhs float32)  {
	_ = this.PropPut(0x00000071, []interface{}{rhs})
}

func (this *Cell) RightPadding() float32 {
	retVal, _ := this.PropGet(0x00000072, nil)
	return retVal.FltValVal()
}

func (this *Cell) SetRightPadding(rhs float32)  {
	_ = this.PropPut(0x00000072, []interface{}{rhs})
}

func (this *Cell) ID() string {
	retVal, _ := this.PropGet(0x00000073, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Cell) SetID(rhs string)  {
	_ = this.PropPut(0x00000073, []interface{}{rhs})
}

func (this *Cell) PreferredWidthType() int32 {
	retVal, _ := this.PropGet(0x00000074, nil)
	return retVal.LValVal()
}

func (this *Cell) SetPreferredWidthType(rhs int32)  {
	_ = this.PropPut(0x00000074, []interface{}{rhs})
}

