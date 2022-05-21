package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020965-0000-0000-C000-000000000046
var IID_Variables = syscall.GUID{0x00020965, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Variables struct {
	ole.OleClient
}

func NewVariables(pDisp *win32.IDispatch, addRef bool, scoped bool) *Variables {
	p := &Variables{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func VariablesFromVar(v ole.Variant) *Variables {
	return NewVariables(v.PdispValVal(), false, false)
}

func (this *Variables) IID() *syscall.GUID {
	return &IID_Variables
}

func (this *Variables) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Variables) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Variables) ForEach(action func(item *Variable) bool) {
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
		pItem := (*Variable)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Variables) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Variables) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Variables) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Variables) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Variables) Item(index *ole.Variant) *Variable {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewVariable(retVal.PdispValVal(), false, true)
}

var Variables_Add_OptArgs= []string{
	"Value", 
}

func (this *Variables) Add(name string, optArgs ...interface{}) *Variable {
	optArgs = ole.ProcessOptArgs(Variables_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000007, []interface{}{name}, optArgs...)
	return NewVariable(retVal.PdispValVal(), false, true)
}

