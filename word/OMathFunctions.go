package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 497142A4-16FD-42C6-BC58-15D89345FC21
var IID_OMathFunctions = syscall.GUID{0x497142A4, 0x16FD, 0x42C6, 
	[8]byte{0xBC, 0x58, 0x15, 0xD8, 0x93, 0x45, 0xFC, 0x21}}

type OMathFunctions struct {
	ole.OleClient
}

func NewOMathFunctions(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathFunctions {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathFunctions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathFunctionsFromVar(v ole.Variant) *OMathFunctions {
	return NewOMathFunctions(v.IDispatch(), false, false)
}

func (this *OMathFunctions) IID() *syscall.GUID {
	return &IID_OMathFunctions
}

func (this *OMathFunctions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathFunctions) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OMathFunctions) ForEach(action func(item *OMathFunction) bool) {
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
		pItem := (*OMathFunction)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *OMathFunctions) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathFunctions) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathFunctions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathFunctions) Count() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *OMathFunctions) Item(index int32) *OMathFunction {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewOMathFunction(retVal.IDispatch(), false, true)
}

var OMathFunctions_Add_OptArgs= []string{
	"NumArgs", "NumCols", 
}

func (this *OMathFunctions) Add(range_ *Range, type_ int32, optArgs ...interface{}) *OMathFunction {
	optArgs = ole.ProcessOptArgs(OMathFunctions_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000068, []interface{}{range_, type_}, optArgs...)
	return NewOMathFunction(retVal.IDispatch(), false, true)
}

