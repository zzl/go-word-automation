package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209D1-0000-0000-C000-000000000046
var IID_HangulAndAlphabetExceptions = syscall.GUID{0x000209D1, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HangulAndAlphabetExceptions struct {
	ole.OleClient
}

func NewHangulAndAlphabetExceptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *HangulAndAlphabetExceptions {
	 if pDisp == nil {
		return nil;
	}
	p := &HangulAndAlphabetExceptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HangulAndAlphabetExceptionsFromVar(v ole.Variant) *HangulAndAlphabetExceptions {
	return NewHangulAndAlphabetExceptions(v.IDispatch(), false, false)
}

func (this *HangulAndAlphabetExceptions) IID() *syscall.GUID {
	return &IID_HangulAndAlphabetExceptions
}

func (this *HangulAndAlphabetExceptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HangulAndAlphabetExceptions) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *HangulAndAlphabetExceptions) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HangulAndAlphabetExceptions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *HangulAndAlphabetExceptions) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *HangulAndAlphabetExceptions) ForEach(action func(item *HangulAndAlphabetException) bool) {
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
		pItem := (*HangulAndAlphabetException)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *HangulAndAlphabetExceptions) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *HangulAndAlphabetExceptions) Item(index *ole.Variant) *HangulAndAlphabetException {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewHangulAndAlphabetException(retVal.IDispatch(), false, true)
}

func (this *HangulAndAlphabetExceptions) Add(name string) *HangulAndAlphabetException {
	retVal, _ := this.Call(0x00000065, []interface{}{name})
	return NewHangulAndAlphabetException(retVal.IDispatch(), false, true)
}

