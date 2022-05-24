package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// F2B60A10-DED5-46FB-A914-3C6F4EBB6451
var IID_SmartTagRecognizers = syscall.GUID{0xF2B60A10, 0xDED5, 0x46FB, 
	[8]byte{0xA9, 0x14, 0x3C, 0x6F, 0x4E, 0xBB, 0x64, 0x51}}

type SmartTagRecognizers struct {
	ole.OleClient
}

func NewSmartTagRecognizers(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTagRecognizers {
	 if pDisp == nil {
		return nil;
	}
	p := &SmartTagRecognizers{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagRecognizersFromVar(v ole.Variant) *SmartTagRecognizers {
	return NewSmartTagRecognizers(v.IDispatch(), false, false)
}

func (this *SmartTagRecognizers) IID() *syscall.GUID {
	return &IID_SmartTagRecognizers
}

func (this *SmartTagRecognizers) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTagRecognizers) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SmartTagRecognizers) ForEach(action func(item *SmartTagRecognizer) bool) {
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
		pItem := (*SmartTagRecognizer)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *SmartTagRecognizers) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *SmartTagRecognizers) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SmartTagRecognizers) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SmartTagRecognizers) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SmartTagRecognizers) Item(index *ole.Variant) *SmartTagRecognizer {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewSmartTagRecognizer(retVal.IDispatch(), false, true)
}

func (this *SmartTagRecognizers) ReloadRecognizers()  {
	retVal, _ := this.Call(0x000003eb, nil)
	_= retVal
}

