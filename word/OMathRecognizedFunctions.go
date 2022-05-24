package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 44FEE887-6600-41AB-95A5-DE33C605116C
var IID_OMathRecognizedFunctions = syscall.GUID{0x44FEE887, 0x6600, 0x41AB, 
	[8]byte{0x95, 0xA5, 0xDE, 0x33, 0xC6, 0x05, 0x11, 0x6C}}

type OMathRecognizedFunctions struct {
	ole.OleClient
}

func NewOMathRecognizedFunctions(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathRecognizedFunctions {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathRecognizedFunctions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathRecognizedFunctionsFromVar(v ole.Variant) *OMathRecognizedFunctions {
	return NewOMathRecognizedFunctions(v.IDispatch(), false, false)
}

func (this *OMathRecognizedFunctions) IID() *syscall.GUID {
	return &IID_OMathRecognizedFunctions
}

func (this *OMathRecognizedFunctions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathRecognizedFunctions) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathRecognizedFunctions) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathRecognizedFunctions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathRecognizedFunctions) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OMathRecognizedFunctions) ForEach(action func(item *OMathRecognizedFunction) bool) {
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
		pItem := (*OMathRecognizedFunction)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *OMathRecognizedFunctions) Count() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *OMathRecognizedFunctions) Item(index *ole.Variant) *OMathRecognizedFunction {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewOMathRecognizedFunction(retVal.IDispatch(), false, true)
}

func (this *OMathRecognizedFunctions) Add(name string) *OMathRecognizedFunction {
	retVal, _ := this.Call(0x000000c8, []interface{}{name})
	return NewOMathRecognizedFunction(retVal.IDispatch(), false, true)
}

