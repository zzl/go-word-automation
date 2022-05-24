package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002094A-0000-0000-C000-000000000046
var IID_Cells = syscall.GUID{0x0002094A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Cells struct {
	ole.OleClient
}

func NewCells(pDisp *win32.IDispatch, addRef bool, scoped bool) *Cells {
	 if pDisp == nil {
		return nil;
	}
	p := &Cells{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CellsFromVar(v ole.Variant) *Cells {
	return NewCells(v.IDispatch(), false, false)
}

func (this *Cells) IID() *syscall.GUID {
	return &IID_Cells
}

func (this *Cells) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Cells) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Cells) ForEach(action func(item *Cell) bool) {
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
		pItem := (*Cell)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Cells) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Cells) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Cells) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Cells) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Cells) Width() float32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.FltValVal()
}

func (this *Cells) SetWidth(rhs float32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *Cells) Height() float32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.FltValVal()
}

func (this *Cells) SetHeight(rhs float32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *Cells) HeightRule() int32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Cells) SetHeightRule(rhs int32)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *Cells) VerticalAlignment() int32 {
	retVal, _ := this.PropGet(0x00000450, nil)
	return retVal.LValVal()
}

func (this *Cells) SetVerticalAlignment(rhs int32)  {
	_ = this.PropPut(0x00000450, []interface{}{rhs})
}

func (this *Cells) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Cells) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Cells) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000065, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Cells) Item(index int32) *Cell {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewCell(retVal.IDispatch(), false, true)
}

var Cells_Add_OptArgs= []string{
	"BeforeCell", 
}

func (this *Cells) Add(optArgs ...interface{}) *Cell {
	optArgs = ole.ProcessOptArgs(Cells_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000004, nil, optArgs...)
	return NewCell(retVal.IDispatch(), false, true)
}

var Cells_Delete_OptArgs= []string{
	"ShiftCells", 
}

func (this *Cells) Delete(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Cells_Delete_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000c8, nil, optArgs...)
	_= retVal
}

func (this *Cells) SetWidth_(columnWidth float32, rulerStyle int32)  {
	retVal, _ := this.Call(0x000000ca, []interface{}{columnWidth, rulerStyle})
	_= retVal
}

func (this *Cells) SetHeight_(rowHeight *ole.Variant, heightRule int32)  {
	retVal, _ := this.Call(0x000000cb, []interface{}{rowHeight, heightRule})
	_= retVal
}

func (this *Cells) Merge()  {
	retVal, _ := this.Call(0x000000cc, nil)
	_= retVal
}

var Cells_Split_OptArgs= []string{
	"NumRows", "NumColumns", "MergeBeforeSplit", 
}

func (this *Cells) Split(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Cells_Split_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000cd, nil, optArgs...)
	_= retVal
}

func (this *Cells) DistributeHeight()  {
	retVal, _ := this.Call(0x000000ce, nil)
	_= retVal
}

func (this *Cells) DistributeWidth()  {
	retVal, _ := this.Call(0x000000cf, nil)
	_= retVal
}

func (this *Cells) AutoFit()  {
	retVal, _ := this.Call(0x000000d0, nil)
	_= retVal
}

func (this *Cells) NestingLevel() int32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *Cells) PreferredWidth() float32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *Cells) SetPreferredWidth(rhs float32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Cells) PreferredWidthType() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *Cells) SetPreferredWidthType(rhs int32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

