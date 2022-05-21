package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209DF-0000-0000-C000-000000000046
var IID_OtherCorrectionsExceptions = syscall.GUID{0x000209DF, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OtherCorrectionsExceptions struct {
	ole.OleClient
}

func NewOtherCorrectionsExceptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *OtherCorrectionsExceptions {
	p := &OtherCorrectionsExceptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OtherCorrectionsExceptionsFromVar(v ole.Variant) *OtherCorrectionsExceptions {
	return NewOtherCorrectionsExceptions(v.PdispValVal(), false, false)
}

func (this *OtherCorrectionsExceptions) IID() *syscall.GUID {
	return &IID_OtherCorrectionsExceptions
}

func (this *OtherCorrectionsExceptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OtherCorrectionsExceptions) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OtherCorrectionsExceptions) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *OtherCorrectionsExceptions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OtherCorrectionsExceptions) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OtherCorrectionsExceptions) ForEach(action func(item *OtherCorrectionsException) bool) {
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
		pItem := (*OtherCorrectionsException)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *OtherCorrectionsExceptions) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *OtherCorrectionsExceptions) Item(index *ole.Variant) *OtherCorrectionsException {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewOtherCorrectionsException(retVal.PdispValVal(), false, true)
}

func (this *OtherCorrectionsExceptions) Add(name string) *OtherCorrectionsException {
	retVal := this.Call(0x00000065, []interface{}{name})
	return NewOtherCorrectionsException(retVal.PdispValVal(), false, true)
}

