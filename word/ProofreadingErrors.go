package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209BB-0000-0000-C000-000000000046
var IID_ProofreadingErrors = syscall.GUID{0x000209BB, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ProofreadingErrors struct {
	ole.OleClient
}

func NewProofreadingErrors(pDisp *win32.IDispatch, addRef bool, scoped bool) *ProofreadingErrors {
	p := &ProofreadingErrors{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ProofreadingErrorsFromVar(v ole.Variant) *ProofreadingErrors {
	return NewProofreadingErrors(v.PdispValVal(), false, false)
}

func (this *ProofreadingErrors) IID() *syscall.GUID {
	return &IID_ProofreadingErrors
}

func (this *ProofreadingErrors) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ProofreadingErrors) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ProofreadingErrors) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ProofreadingErrors) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ProofreadingErrors) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ProofreadingErrors) ForEach(action func(item *Range) bool) {
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
		pItem := (*Range)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ProofreadingErrors) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *ProofreadingErrors) Type() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *ProofreadingErrors) Item(index int32) *Range {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewRange(retVal.PdispValVal(), false, true)
}

