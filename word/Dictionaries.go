package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209AC-0000-0000-C000-000000000046
var IID_Dictionaries = syscall.GUID{0x000209AC, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Dictionaries struct {
	ole.OleClient
}

func NewDictionaries(pDisp *win32.IDispatch, addRef bool, scoped bool) *Dictionaries {
	 if pDisp == nil {
		return nil;
	}
	p := &Dictionaries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DictionariesFromVar(v ole.Variant) *Dictionaries {
	return NewDictionaries(v.IDispatch(), false, false)
}

func (this *Dictionaries) IID() *syscall.GUID {
	return &IID_Dictionaries
}

func (this *Dictionaries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Dictionaries) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Dictionaries) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Dictionaries) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Dictionaries) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Dictionaries) ForEach(action func(item *Dictionary) bool) {
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

func (this *Dictionaries) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Dictionaries) Maximum() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Dictionaries) ActiveCustomDictionary() *Dictionary {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *Dictionaries) SetActiveCustomDictionary(rhs *Dictionary)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Dictionaries) Item(index *ole.Variant) *Dictionary {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *Dictionaries) Add(fileName string) *Dictionary {
	retVal, _ := this.Call(0x00000065, []interface{}{fileName})
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *Dictionaries) ClearAll()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

