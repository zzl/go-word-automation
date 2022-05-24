package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// CDE12CD8-767B-4757-8A31-13029A086305
var IID_SmartTagActions = syscall.GUID{0xCDE12CD8, 0x767B, 0x4757, 
	[8]byte{0x8A, 0x31, 0x13, 0x02, 0x9A, 0x08, 0x63, 0x05}}

type SmartTagActions struct {
	ole.OleClient
}

func NewSmartTagActions(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTagActions {
	 if pDisp == nil {
		return nil;
	}
	p := &SmartTagActions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagActionsFromVar(v ole.Variant) *SmartTagActions {
	return NewSmartTagActions(v.IDispatch(), false, false)
}

func (this *SmartTagActions) IID() *syscall.GUID {
	return &IID_SmartTagActions
}

func (this *SmartTagActions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTagActions) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SmartTagActions) ForEach(action func(item *SmartTagAction) bool) {
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
		pItem := (*SmartTagAction)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *SmartTagActions) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *SmartTagActions) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SmartTagActions) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SmartTagActions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SmartTagActions) Item(index *ole.Variant) *SmartTagAction {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewSmartTagAction(retVal.IDispatch(), false, true)
}

func (this *SmartTagActions) ReloadActions()  {
	retVal, _ := this.Call(0x000003eb, nil)
	_= retVal
}

