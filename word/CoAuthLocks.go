package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// DFF99AC2-CD2A-43AD-91B1-A2BE40BC7146
var IID_CoAuthLocks = syscall.GUID{0xDFF99AC2, 0xCD2A, 0x43AD, 
	[8]byte{0x91, 0xB1, 0xA2, 0xBE, 0x40, 0xBC, 0x71, 0x46}}

type CoAuthLocks struct {
	ole.OleClient
}

func NewCoAuthLocks(pDisp *win32.IDispatch, addRef bool, scoped bool) *CoAuthLocks {
	 if pDisp == nil {
		return nil;
	}
	p := &CoAuthLocks{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CoAuthLocksFromVar(v ole.Variant) *CoAuthLocks {
	return NewCoAuthLocks(v.IDispatch(), false, false)
}

func (this *CoAuthLocks) IID() *syscall.GUID {
	return &IID_CoAuthLocks
}

func (this *CoAuthLocks) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CoAuthLocks) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CoAuthLocks) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CoAuthLocks) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CoAuthLocks) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *CoAuthLocks) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CoAuthLocks) ForEach(action func(item *CoAuthLock) bool) {
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
		pItem := (*CoAuthLock)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *CoAuthLocks) Item(index int32) *CoAuthLock {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewCoAuthLock(retVal.IDispatch(), false, true)
}

var CoAuthLocks_Add_OptArgs= []string{
	"Range", "Type", 
}

func (this *CoAuthLocks) Add(optArgs ...interface{}) *CoAuthLock {
	optArgs = ole.ProcessOptArgs(CoAuthLocks_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000002, nil, optArgs...)
	return NewCoAuthLock(retVal.IDispatch(), false, true)
}

func (this *CoAuthLocks) RemoveEphemeralLocks()  {
	retVal, _ := this.Call(0x00000003, nil)
	_= retVal
}

