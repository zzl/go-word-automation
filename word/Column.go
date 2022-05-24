package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002094F-0000-0000-C000-000000000046
var IID_Column = syscall.GUID{0x0002094F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Column struct {
	ole.OleClient
}

func NewColumn(pDisp *win32.IDispatch, addRef bool, scoped bool) *Column {
	 if pDisp == nil {
		return nil;
	}
	p := &Column{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ColumnFromVar(v ole.Variant) *Column {
	return NewColumn(v.IDispatch(), false, false)
}

func (this *Column) IID() *syscall.GUID {
	return &IID_Column
}

func (this *Column) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Column) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Column) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Column) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Column) Width() float32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.FltValVal()
}

func (this *Column) SetWidth(rhs float32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Column) IsFirst() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Column) IsLast() bool {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Column) Index() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Column) Cells() *Cells {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewCells(retVal.IDispatch(), false, true)
}

func (this *Column) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Column) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Column) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000066, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Column) Next() *Column {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewColumn(retVal.IDispatch(), false, true)
}

func (this *Column) Previous() *Column {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewColumn(retVal.IDispatch(), false, true)
}

func (this *Column) Select()  {
	retVal, _ := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *Column) Delete()  {
	retVal, _ := this.Call(0x000000c8, nil)
	_= retVal
}

func (this *Column) SetWidth_(columnWidth float32, rulerStyle int32)  {
	retVal, _ := this.Call(0x000000c9, []interface{}{columnWidth, rulerStyle})
	_= retVal
}

func (this *Column) AutoFit()  {
	retVal, _ := this.Call(0x000000ca, nil)
	_= retVal
}

var Column_SortOld_OptArgs= []string{
	"ExcludeHeader", "SortFieldType", "SortOrder", "CaseSensitive", "LanguageID", 
}

func (this *Column) SortOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Column_SortOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000cb, nil, optArgs...)
	_= retVal
}

var Column_Sort_OptArgs= []string{
	"ExcludeHeader", "SortFieldType", "SortOrder", "CaseSensitive", 
	"BidiSort", "IgnoreThe", "IgnoreKashida", "IgnoreDiacritics", 
	"IgnoreHe", "LanguageID", 
}

func (this *Column) Sort(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Column_Sort_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000cc, nil, optArgs...)
	_= retVal
}

func (this *Column) NestingLevel() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *Column) PreferredWidth() float32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.FltValVal()
}

func (this *Column) SetPreferredWidth(rhs float32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *Column) PreferredWidthType() int32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *Column) SetPreferredWidthType(rhs int32)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

