package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020961-0000-0000-C000-000000000046
var IID_Windows = syscall.GUID{0x00020961, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Windows struct {
	ole.OleClient
}

func NewWindows(pDisp *win32.IDispatch, addRef bool, scoped bool) *Windows {
	 if pDisp == nil {
		return nil;
	}
	p := &Windows{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WindowsFromVar(v ole.Variant) *Windows {
	return NewWindows(v.IDispatch(), false, false)
}

func (this *Windows) IID() *syscall.GUID {
	return &IID_Windows
}

func (this *Windows) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Windows) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Windows) ForEach(action func(item *Window) bool) {
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
		pItem := (*Window)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Windows) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Windows) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Windows) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Windows) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Windows) Item(index *ole.Variant) *Window {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewWindow(retVal.IDispatch(), false, true)
}

var Windows_Add_OptArgs= []string{
	"Window", 
}

func (this *Windows) Add(optArgs ...interface{}) *Window {
	optArgs = ole.ProcessOptArgs(Windows_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000a, nil, optArgs...)
	return NewWindow(retVal.IDispatch(), false, true)
}

var Windows_Arrange_OptArgs= []string{
	"ArrangeStyle", 
}

func (this *Windows) Arrange(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Windows_Arrange_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000b, nil, optArgs...)
	_= retVal
}

func (this *Windows) CompareSideBySideWith(document *ole.Variant) bool {
	retVal, _ := this.Call(0x0000000c, []interface{}{document})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Windows) BreakSideBySide() bool {
	retVal, _ := this.Call(0x0000000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Windows) ResetPositionsSideBySide()  {
	retVal, _ := this.Call(0x0000000e, nil)
	_= retVal
}

func (this *Windows) SyncScrollingSideBySide() bool {
	retVal, _ := this.PropGet(0x000003eb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Windows) SetSyncScrollingSideBySide(rhs bool)  {
	_ = this.PropPut(0x000003eb, []interface{}{rhs})
}

