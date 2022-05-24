package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209E0-0000-0000-C000-000000000046
var IID_HangulHanjaConversionDictionaries = syscall.GUID{0x000209E0, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HangulHanjaConversionDictionaries struct {
	ole.OleClient
}

func NewHangulHanjaConversionDictionaries(pDisp *win32.IDispatch, addRef bool, scoped bool) *HangulHanjaConversionDictionaries {
	 if pDisp == nil {
		return nil;
	}
	p := &HangulHanjaConversionDictionaries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HangulHanjaConversionDictionariesFromVar(v ole.Variant) *HangulHanjaConversionDictionaries {
	return NewHangulHanjaConversionDictionaries(v.IDispatch(), false, false)
}

func (this *HangulHanjaConversionDictionaries) IID() *syscall.GUID {
	return &IID_HangulHanjaConversionDictionaries
}

func (this *HangulHanjaConversionDictionaries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HangulHanjaConversionDictionaries) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *HangulHanjaConversionDictionaries) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HangulHanjaConversionDictionaries) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *HangulHanjaConversionDictionaries) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *HangulHanjaConversionDictionaries) ForEach(action func(item *Dictionary) bool) {
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
		pItem := (*Dictionary)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *HangulHanjaConversionDictionaries) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *HangulHanjaConversionDictionaries) Maximum() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *HangulHanjaConversionDictionaries) ActiveCustomDictionary() *Dictionary {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *HangulHanjaConversionDictionaries) SetActiveCustomDictionary(rhs *Dictionary)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *HangulHanjaConversionDictionaries) BuiltinDictionary() *Dictionary {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *HangulHanjaConversionDictionaries) Item(index *ole.Variant) *Dictionary {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *HangulHanjaConversionDictionaries) Add(fileName string) *Dictionary {
	retVal, _ := this.Call(0x00000065, []interface{}{fileName})
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *HangulHanjaConversionDictionaries) ClearAll()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

