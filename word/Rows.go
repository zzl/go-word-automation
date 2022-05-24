package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002094C-0000-0000-C000-000000000046
var IID_Rows = syscall.GUID{0x0002094C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Rows struct {
	ole.OleClient
}

func NewRows(pDisp *win32.IDispatch, addRef bool, scoped bool) *Rows {
	 if pDisp == nil {
		return nil;
	}
	p := &Rows{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RowsFromVar(v ole.Variant) *Rows {
	return NewRows(v.IDispatch(), false, false)
}

func (this *Rows) IID() *syscall.GUID {
	return &IID_Rows
}

func (this *Rows) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Rows) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Rows) ForEach(action func(item *Row) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*Row)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Rows) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Rows) AllowBreakAcrossPages() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Rows) SetAllowBreakAcrossPages(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Rows) Alignment() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Rows) SetAlignment(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *Rows) HeadingFormat() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Rows) SetHeadingFormat(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *Rows) SpaceBetweenColumns() float32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.FltValVal()
}

func (this *Rows) SetSpaceBetweenColumns(rhs float32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *Rows) Height() float32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.FltValVal()
}

func (this *Rows) SetHeight(rhs float32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *Rows) HeightRule() int32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Rows) SetHeightRule(rhs int32)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *Rows) LeftIndent() float32 {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.FltValVal()
}

func (this *Rows) SetLeftIndent(rhs float32)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *Rows) First() *Row {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewRow(retVal.IDispatch(), false, true)
}

func (this *Rows) Last() *Row {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return NewRow(retVal.IDispatch(), false, true)
}

func (this *Rows) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Rows) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Rows) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Rows) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Rows) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Rows) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000066, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Rows) Item(index int32) *Row {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewRow(retVal.IDispatch(), false, true)
}

var Rows_Add_OptArgs= []string{
	"BeforeRow", 
}

func (this *Rows) Add(optArgs ...interface{}) *Row {
	optArgs = ole.ProcessOptArgs(Rows_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000064, nil, optArgs...)
	return NewRow(retVal.IDispatch(), false, true)
}

func (this *Rows) Select()  {
	retVal, _ := this.Call(0x000000c7, nil)
	_= retVal
}

func (this *Rows) Delete()  {
	retVal, _ := this.Call(0x000000c8, nil)
	_= retVal
}

func (this *Rows) SetLeftIndent_(leftIndent float32, rulerStyle int32)  {
	retVal, _ := this.Call(0x000000ca, []interface{}{leftIndent, rulerStyle})
	_= retVal
}

func (this *Rows) SetHeight_(rowHeight float32, heightRule int32)  {
	retVal, _ := this.Call(0x000000cb, []interface{}{rowHeight, heightRule})
	_= retVal
}

var Rows_ConvertToTextOld_OptArgs= []string{
	"Separator", 
}

func (this *Rows) ConvertToTextOld(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Rows_ConvertToTextOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000010, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Rows) DistributeHeight()  {
	retVal, _ := this.Call(0x000000ce, nil)
	_= retVal
}

var Rows_ConvertToText_OptArgs= []string{
	"Separator", "NestedTables", 
}

func (this *Rows) ConvertToText(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Rows_ConvertToText_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d2, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Rows) WrapAroundText() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *Rows) SetWrapAroundText(rhs int32)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *Rows) DistanceTop() float32 {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.FltValVal()
}

func (this *Rows) SetDistanceTop(rhs float32)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *Rows) DistanceBottom() float32 {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return retVal.FltValVal()
}

func (this *Rows) SetDistanceBottom(rhs float32)  {
	_ = this.PropPut(0x0000000e, []interface{}{rhs})
}

func (this *Rows) DistanceLeft() float32 {
	retVal, _ := this.PropGet(0x00000014, nil)
	return retVal.FltValVal()
}

func (this *Rows) SetDistanceLeft(rhs float32)  {
	_ = this.PropPut(0x00000014, []interface{}{rhs})
}

func (this *Rows) DistanceRight() float32 {
	retVal, _ := this.PropGet(0x00000015, nil)
	return retVal.FltValVal()
}

func (this *Rows) SetDistanceRight(rhs float32)  {
	_ = this.PropPut(0x00000015, []interface{}{rhs})
}

func (this *Rows) HorizontalPosition() float32 {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return retVal.FltValVal()
}

func (this *Rows) SetHorizontalPosition(rhs float32)  {
	_ = this.PropPut(0x0000000f, []interface{}{rhs})
}

func (this *Rows) VerticalPosition() float32 {
	retVal, _ := this.PropGet(0x00000011, nil)
	return retVal.FltValVal()
}

func (this *Rows) SetVerticalPosition(rhs float32)  {
	_ = this.PropPut(0x00000011, []interface{}{rhs})
}

func (this *Rows) RelativeHorizontalPosition() int32 {
	retVal, _ := this.PropGet(0x00000012, nil)
	return retVal.LValVal()
}

func (this *Rows) SetRelativeHorizontalPosition(rhs int32)  {
	_ = this.PropPut(0x00000012, []interface{}{rhs})
}

func (this *Rows) RelativeVerticalPosition() int32 {
	retVal, _ := this.PropGet(0x00000013, nil)
	return retVal.LValVal()
}

func (this *Rows) SetRelativeVerticalPosition(rhs int32)  {
	_ = this.PropPut(0x00000013, []interface{}{rhs})
}

func (this *Rows) AllowOverlap() int32 {
	retVal, _ := this.PropGet(0x00000016, nil)
	return retVal.LValVal()
}

func (this *Rows) SetAllowOverlap(rhs int32)  {
	_ = this.PropPut(0x00000016, []interface{}{rhs})
}

func (this *Rows) NestingLevel() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *Rows) TableDirection() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *Rows) SetTableDirection(rhs int32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

