package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020916-0000-0000-C000-000000000046
var IID_CustomLabels = syscall.GUID{0x00020916, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CustomLabels struct {
	ole.OleClient
}

func NewCustomLabels(pDisp *win32.IDispatch, addRef bool, scoped bool) *CustomLabels {
	p := &CustomLabels{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CustomLabelsFromVar(v ole.Variant) *CustomLabels {
	return NewCustomLabels(v.PdispValVal(), false, false)
}

func (this *CustomLabels) IID() *syscall.GUID {
	return &IID_CustomLabels
}

func (this *CustomLabels) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CustomLabels) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CustomLabels) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CustomLabels) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CustomLabels) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CustomLabels) ForEach(action func(item *CustomLabel) bool) {
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
		pItem := (*CustomLabel)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *CustomLabels) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *CustomLabels) Item(index *ole.Variant) *CustomLabel {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewCustomLabel(retVal.PdispValVal(), false, true)
}

var CustomLabels_Add_OptArgs= []string{
	"DotMatrix", 
}

func (this *CustomLabels) Add(name string, optArgs ...interface{}) *CustomLabel {
	optArgs = ole.ProcessOptArgs(CustomLabels_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000065, []interface{}{name}, optArgs...)
	return NewCustomLabel(retVal.PdispValVal(), false, true)
}

