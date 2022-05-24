package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// DB8E3072-E106-4453-8E7C-53056F1D30DC
var IID_SmartTagTypes = syscall.GUID{0xDB8E3072, 0xE106, 0x4453, 
	[8]byte{0x8E, 0x7C, 0x53, 0x05, 0x6F, 0x1D, 0x30, 0xDC}}

type SmartTagTypes struct {
	ole.OleClient
}

func NewSmartTagTypes(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTagTypes {
	 if pDisp == nil {
		return nil;
	}
	p := &SmartTagTypes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagTypesFromVar(v ole.Variant) *SmartTagTypes {
	return NewSmartTagTypes(v.IDispatch(), false, false)
}

func (this *SmartTagTypes) IID() *syscall.GUID {
	return &IID_SmartTagTypes
}

func (this *SmartTagTypes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTagTypes) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SmartTagTypes) ForEach(action func(item *SmartTagType) bool) {
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
		pItem := (*SmartTagType)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *SmartTagTypes) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *SmartTagTypes) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SmartTagTypes) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SmartTagTypes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SmartTagTypes) Item(index *ole.Variant) *SmartTagType {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewSmartTagType(retVal.IDispatch(), false, true)
}

func (this *SmartTagTypes) ReloadAll()  {
	retVal, _ := this.Call(0x000003eb, nil)
	_= retVal
}

