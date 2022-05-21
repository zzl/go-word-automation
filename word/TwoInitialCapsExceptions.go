package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020944-0000-0000-C000-000000000046
var IID_TwoInitialCapsExceptions = syscall.GUID{0x00020944, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TwoInitialCapsExceptions struct {
	ole.OleClient
}

func NewTwoInitialCapsExceptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *TwoInitialCapsExceptions {
	p := &TwoInitialCapsExceptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TwoInitialCapsExceptionsFromVar(v ole.Variant) *TwoInitialCapsExceptions {
	return NewTwoInitialCapsExceptions(v.PdispValVal(), false, false)
}

func (this *TwoInitialCapsExceptions) IID() *syscall.GUID {
	return &IID_TwoInitialCapsExceptions
}

func (this *TwoInitialCapsExceptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TwoInitialCapsExceptions) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *TwoInitialCapsExceptions) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TwoInitialCapsExceptions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *TwoInitialCapsExceptions) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *TwoInitialCapsExceptions) ForEach(action func(item *TwoInitialCapsException) bool) {
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
		pItem := (*TwoInitialCapsException)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *TwoInitialCapsExceptions) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *TwoInitialCapsExceptions) Item(index *ole.Variant) *TwoInitialCapsException {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewTwoInitialCapsException(retVal.PdispValVal(), false, true)
}

func (this *TwoInitialCapsExceptions) Add(name string) *TwoInitialCapsException {
	retVal := this.Call(0x00000065, []interface{}{name})
	return NewTwoInitialCapsException(retVal.PdispValVal(), false, true)
}

