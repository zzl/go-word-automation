package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// C2B83A65-B061-4469-83B6-8877437CB8A0
var IID_Conflicts = syscall.GUID{0xC2B83A65, 0xB061, 0x4469, 
	[8]byte{0x83, 0xB6, 0x88, 0x77, 0x43, 0x7C, 0xB8, 0xA0}}

type Conflicts struct {
	ole.OleClient
}

func NewConflicts(pDisp *win32.IDispatch, addRef bool, scoped bool) *Conflicts {
	 if pDisp == nil {
		return nil;
	}
	p := &Conflicts{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ConflictsFromVar(v ole.Variant) *Conflicts {
	return NewConflicts(v.IDispatch(), false, false)
}

func (this *Conflicts) IID() *syscall.GUID {
	return &IID_Conflicts
}

func (this *Conflicts) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Conflicts) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Conflicts) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Conflicts) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Conflicts) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Conflicts) ForEach(action func(item *Conflict) bool) {
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
		pItem := (*Conflict)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Conflicts) Count() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Conflicts) Item(index int32) *Conflict {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewConflict(retVal.IDispatch(), false, true)
}

func (this *Conflicts) AcceptAll()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *Conflicts) RejectAll()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

