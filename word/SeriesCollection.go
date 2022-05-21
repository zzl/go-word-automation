package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 8FEB78F7-35C6-4871-918C-193C3CDD886D
var IID_SeriesCollection = syscall.GUID{0x8FEB78F7, 0x35C6, 0x4871, 
	[8]byte{0x91, 0x8C, 0x19, 0x3C, 0x3C, 0xDD, 0x88, 0x6D}}

type SeriesCollection struct {
	ole.OleClient
}

func NewSeriesCollection(pDisp *win32.IDispatch, addRef bool, scoped bool) *SeriesCollection {
	p := &SeriesCollection{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SeriesCollectionFromVar(v ole.Variant) *SeriesCollection {
	return NewSeriesCollection(v.PdispValVal(), false, false)
}

func (this *SeriesCollection) IID() *syscall.GUID {
	return &IID_SeriesCollection
}

func (this *SeriesCollection) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SeriesCollection) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

var SeriesCollection_Add_OptArgs= []string{
	"SeriesLabels", "CategoryLabels", "Replace", 
}

func (this *SeriesCollection) Add(source interface{}, rowcol int32, optArgs ...interface{}) *Series {
	optArgs = ole.ProcessOptArgs(SeriesCollection_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{source, rowcol}, optArgs...)
	return NewSeries(retVal.PdispValVal(), false, true)
}

func (this *SeriesCollection) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var SeriesCollection_Extend_OptArgs= []string{
	"Rowcol", "CategoryLabels", 
}

func (this *SeriesCollection) Extend(source interface{}, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(SeriesCollection_Extend_OptArgs, optArgs)
	retVal := this.Call(0x000000e3, []interface{}{source}, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *SeriesCollection) Item(index interface{}) *Series {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewSeries(retVal.PdispValVal(), false, true)
}

func (this *SeriesCollection) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SeriesCollection) ForEach(action func(item *Series) bool) {
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
		pItem := (*Series)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *SeriesCollection) NewSeries() *Series {
	retVal := this.Call(0x0000045d, nil)
	return NewSeries(retVal.PdispValVal(), false, true)
}

func (this *SeriesCollection) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SeriesCollection) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *SeriesCollection) Default_(index interface{}) *Series {
	retVal := this.Call(0x6002000a, []interface{}{index})
	return NewSeries(retVal.PdispValVal(), false, true)
}

