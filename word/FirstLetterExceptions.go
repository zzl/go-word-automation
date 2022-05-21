package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020946-0000-0000-C000-000000000046
var IID_FirstLetterExceptions = syscall.GUID{0x00020946, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FirstLetterExceptions struct {
	ole.OleClient
}

func NewFirstLetterExceptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *FirstLetterExceptions {
	p := &FirstLetterExceptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FirstLetterExceptionsFromVar(v ole.Variant) *FirstLetterExceptions {
	return NewFirstLetterExceptions(v.PdispValVal(), false, false)
}

func (this *FirstLetterExceptions) IID() *syscall.GUID {
	return &IID_FirstLetterExceptions
}

func (this *FirstLetterExceptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FirstLetterExceptions) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *FirstLetterExceptions) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *FirstLetterExceptions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FirstLetterExceptions) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *FirstLetterExceptions) ForEach(action func(item *FirstLetterException) bool) {
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
		pItem := (*FirstLetterException)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *FirstLetterExceptions) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *FirstLetterExceptions) Item(index *ole.Variant) *FirstLetterException {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewFirstLetterException(retVal.PdispValVal(), false, true)
}

func (this *FirstLetterExceptions) Add(name string) *FirstLetterException {
	retVal := this.Call(0x00000065, []interface{}{name})
	return NewFirstLetterException(retVal.PdispValVal(), false, true)
}

