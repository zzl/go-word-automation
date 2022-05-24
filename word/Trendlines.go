package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 54B7061A-D56C-40E5-B85B-58146446C782
var IID_Trendlines = syscall.GUID{0x54B7061A, 0xD56C, 0x40E5, 
	[8]byte{0xB8, 0x5B, 0x58, 0x14, 0x64, 0x46, 0xC7, 0x82}}

type Trendlines struct {
	ole.OleClient
}

func NewTrendlines(pDisp *win32.IDispatch, addRef bool, scoped bool) *Trendlines {
	 if pDisp == nil {
		return nil;
	}
	p := &Trendlines{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TrendlinesFromVar(v ole.Variant) *Trendlines {
	return NewTrendlines(v.IDispatch(), false, false)
}

func (this *Trendlines) IID() *syscall.GUID {
	return &IID_Trendlines
}

func (this *Trendlines) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Trendlines) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Trendlines_Add_OptArgs= []string{
	"Type", "Order", "Period", "Forward", 
	"Backward", "Intercept", "DisplayEquation", "DisplayRSquared", "Name", 
}

func (this *Trendlines) Add(optArgs ...interface{}) *Trendline {
	optArgs = ole.ProcessOptArgs(Trendlines_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b5, nil, optArgs...)
	return NewTrendline(retVal.IDispatch(), false, true)
}

func (this *Trendlines) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

var Trendlines_Item_OptArgs= []string{
	"Index", 
}

func (this *Trendlines) Item(optArgs ...interface{}) *Trendline {
	optArgs = ole.ProcessOptArgs(Trendlines_Item_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000000, nil, optArgs...)
	return NewTrendline(retVal.IDispatch(), false, true)
}

func (this *Trendlines) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Trendlines) ForEach(action func(item *Trendline) bool) {
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
		pItem := (*Trendline)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Trendlines) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Trendlines) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

var Trendlines_Default__OptArgs= []string{
	"Index", 
}

func (this *Trendlines) Default_(optArgs ...interface{}) *Trendline {
	optArgs = ole.ProcessOptArgs(Trendlines_Default__OptArgs, optArgs)
	retVal, _ := this.Call(0x60020007, nil, optArgs...)
	return NewTrendline(retVal.IDispatch(), false, true)
}

