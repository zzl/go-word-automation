package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002092D-0000-0000-C000-000000000046
var IID_Styles = syscall.GUID{0x0002092D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Styles struct {
	ole.OleClient
}

func NewStyles(pDisp *win32.IDispatch, addRef bool, scoped bool) *Styles {
	 if pDisp == nil {
		return nil;
	}
	p := &Styles{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func StylesFromVar(v ole.Variant) *Styles {
	return NewStyles(v.IDispatch(), false, false)
}

func (this *Styles) IID() *syscall.GUID {
	return &IID_Styles
}

func (this *Styles) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Styles) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Styles) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Styles) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Styles) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Styles) ForEach(action func(item *Style) bool) {
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
		pItem := (*Style)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Styles) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Styles) Item(index *ole.Variant) *Style {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewStyle(retVal.IDispatch(), false, true)
}

var Styles_Add_OptArgs= []string{
	"Type", 
}

func (this *Styles) Add(name string, optArgs ...interface{}) *Style {
	optArgs = ole.ProcessOptArgs(Styles_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000064, []interface{}{name}, optArgs...)
	return NewStyle(retVal.IDispatch(), false, true)
}

