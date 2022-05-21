package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020955-0000-0000-C000-000000000046
var IID_TabStops = syscall.GUID{0x00020955, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TabStops struct {
	ole.OleClient
}

func NewTabStops(pDisp *win32.IDispatch, addRef bool, scoped bool) *TabStops {
	p := &TabStops{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TabStopsFromVar(v ole.Variant) *TabStops {
	return NewTabStops(v.PdispValVal(), false, false)
}

func (this *TabStops) IID() *syscall.GUID {
	return &IID_TabStops
}

func (this *TabStops) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TabStops) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *TabStops) ForEach(action func(item *TabStop) bool) {
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
		pItem := (*TabStop)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *TabStops) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *TabStops) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *TabStops) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TabStops) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *TabStops) Item(index *ole.Variant) *TabStop {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewTabStop(retVal.PdispValVal(), false, true)
}

var TabStops_Add_OptArgs= []string{
	"Alignment", "Leader", 
}

func (this *TabStops) Add(position float32, optArgs ...interface{}) *TabStop {
	optArgs = ole.ProcessOptArgs(TabStops_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000064, []interface{}{position}, optArgs...)
	return NewTabStop(retVal.PdispValVal(), false, true)
}

func (this *TabStops) ClearAll()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *TabStops) Before(position float32) *TabStop {
	retVal := this.Call(0x00000066, []interface{}{position})
	return NewTabStop(retVal.PdispValVal(), false, true)
}

func (this *TabStops) After(position float32) *TabStop {
	retVal := this.Call(0x00000067, []interface{}{position})
	return NewTabStop(retVal.PdispValVal(), false, true)
}

