package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 30225CFC-5A71-4FE6-B527-90A52C54AE77
var IID_CoAuthUpdates = syscall.GUID{0x30225CFC, 0x5A71, 0x4FE6, 
	[8]byte{0xB5, 0x27, 0x90, 0xA5, 0x2C, 0x54, 0xAE, 0x77}}

type CoAuthUpdates struct {
	ole.OleClient
}

func NewCoAuthUpdates(pDisp *win32.IDispatch, addRef bool, scoped bool) *CoAuthUpdates {
	p := &CoAuthUpdates{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CoAuthUpdatesFromVar(v ole.Variant) *CoAuthUpdates {
	return NewCoAuthUpdates(v.PdispValVal(), false, false)
}

func (this *CoAuthUpdates) IID() *syscall.GUID {
	return &IID_CoAuthUpdates
}

func (this *CoAuthUpdates) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CoAuthUpdates) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CoAuthUpdates) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CoAuthUpdates) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CoAuthUpdates) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *CoAuthUpdates) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CoAuthUpdates) ForEach(action func(item *CoAuthUpdate) bool) {
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
		pItem := (*CoAuthUpdate)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *CoAuthUpdates) Item(index int32) *CoAuthUpdate {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewCoAuthUpdate(retVal.PdispValVal(), false, true)
}

