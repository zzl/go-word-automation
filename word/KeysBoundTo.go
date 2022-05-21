package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020997-0000-0000-C000-000000000046
var IID_KeysBoundTo = syscall.GUID{0x00020997, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type KeysBoundTo struct {
	ole.OleClient
}

func NewKeysBoundTo(pDisp *win32.IDispatch, addRef bool, scoped bool) *KeysBoundTo {
	p := &KeysBoundTo{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func KeysBoundToFromVar(v ole.Variant) *KeysBoundTo {
	return NewKeysBoundTo(v.PdispValVal(), false, false)
}

func (this *KeysBoundTo) IID() *syscall.GUID {
	return &IID_KeysBoundTo
}

func (this *KeysBoundTo) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *KeysBoundTo) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *KeysBoundTo) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *KeysBoundTo) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *KeysBoundTo) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *KeysBoundTo) ForEach(action func(item *KeyBinding) bool) {
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
		pItem := (*KeyBinding)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *KeysBoundTo) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *KeysBoundTo) KeyCategory() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *KeysBoundTo) Command() string {
	retVal := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *KeysBoundTo) CommandParameter() string {
	retVal := this.PropGet(0x00000005, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *KeysBoundTo) Context() *ole.DispatchClass {
	retVal := this.PropGet(0x0000000a, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *KeysBoundTo) Item(index int32) *KeyBinding {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewKeyBinding(retVal.PdispValVal(), false, true)
}

var KeysBoundTo_Key_OptArgs= []string{
	"KeyCode2", 
}

func (this *KeysBoundTo) Key(keyCode int32, optArgs ...interface{}) *KeyBinding {
	optArgs = ole.ProcessOptArgs(KeysBoundTo_Key_OptArgs, optArgs)
	retVal := this.Call(0x00000001, []interface{}{keyCode}, optArgs...)
	return NewKeyBinding(retVal.PdispValVal(), false, true)
}

