package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002094B-0000-0000-C000-000000000046
var IID_Columns = syscall.GUID{0x0002094B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Columns struct {
	ole.OleClient
}

func NewColumns(pDisp *win32.IDispatch, addRef bool, scoped bool) *Columns {
	 if pDisp == nil {
		return nil;
	}
	p := &Columns{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ColumnsFromVar(v ole.Variant) *Columns {
	return NewColumns(v.IDispatch(), false, false)
}

func (this *Columns) IID() *syscall.GUID {
	return &IID_Columns
}

func (this *Columns) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Columns) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Columns) ForEach(action func(item *Column) bool) {
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
		pItem := (*Column)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Columns) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Columns) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Columns) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Columns) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Columns) First() *Column {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewColumn(retVal.IDispatch(), false, true)
}

func (this *Columns) Last() *Column {
	retVal, _ := this.PropGet(0x00000065, nil)
	return NewColumn(retVal.IDispatch(), false, true)
}

func (this *Columns) Width() float32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.FltValVal()
}

func (this *Columns) SetWidth(rhs float32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Columns) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Columns) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Columns) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Columns) Item(index int32) *Column {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewColumn(retVal.IDispatch(), false, true)
}

var Columns_Add_OptArgs= []string{
	"BeforeColumn", 
}

func (this *Columns) Add(optArgs ...interface{}) *Column {
	optArgs = ole.ProcessOptArgs(Columns_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000005, nil, optArgs...)
	return NewColumn(retVal.IDispatch(), false, true)
}

func (this *Columns) Select()  {
	retVal, _ := this.Call(0x000000c7, nil)
	_= retVal
}

func (this *Columns) Delete()  {
	retVal, _ := this.Call(0x000000c8, nil)
	_= retVal
}

func (this *Columns) SetWidth_(columnWidth float32, rulerStyle int32)  {
	retVal, _ := this.Call(0x000000c9, []interface{}{columnWidth, rulerStyle})
	_= retVal
}

func (this *Columns) AutoFit()  {
	retVal, _ := this.Call(0x000000ca, nil)
	_= retVal
}

func (this *Columns) DistributeWidth()  {
	retVal, _ := this.Call(0x000000cb, nil)
	_= retVal
}

func (this *Columns) NestingLevel() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *Columns) PreferredWidth() float32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.FltValVal()
}

func (this *Columns) SetPreferredWidth(rhs float32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *Columns) PreferredWidthType() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *Columns) SetPreferredWidthType(rhs int32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

