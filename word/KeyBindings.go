package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020996-0000-0000-C000-000000000046
var IID_KeyBindings = syscall.GUID{0x00020996, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type KeyBindings struct {
	ole.OleClient
}

func NewKeyBindings(pDisp *win32.IDispatch, addRef bool, scoped bool) *KeyBindings {
	 if pDisp == nil {
		return nil;
	}
	p := &KeyBindings{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func KeyBindingsFromVar(v ole.Variant) *KeyBindings {
	return NewKeyBindings(v.IDispatch(), false, false)
}

func (this *KeyBindings) IID() *syscall.GUID {
	return &IID_KeyBindings
}

func (this *KeyBindings) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *KeyBindings) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *KeyBindings) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *KeyBindings) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *KeyBindings) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *KeyBindings) ForEach(action func(item *KeyBinding) bool) {
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

func (this *KeyBindings) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *KeyBindings) Context() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *KeyBindings) Item(index int32) *KeyBinding {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewKeyBinding(retVal.IDispatch(), false, true)
}

var KeyBindings_Add_OptArgs= []string{
	"KeyCode2", "CommandParameter", 
}

func (this *KeyBindings) Add(keyCategory int32, command string, keyCode int32, optArgs ...interface{}) *KeyBinding {
	optArgs = ole.ProcessOptArgs(KeyBindings_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, []interface{}{keyCategory, command, keyCode}, optArgs...)
	return NewKeyBinding(retVal.IDispatch(), false, true)
}

func (this *KeyBindings) ClearAll()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

var KeyBindings_Key_OptArgs= []string{
	"KeyCode2", 
}

func (this *KeyBindings) Key(keyCode int32, optArgs ...interface{}) *KeyBinding {
	optArgs = ole.ProcessOptArgs(KeyBindings_Key_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006e, []interface{}{keyCode}, optArgs...)
	return NewKeyBinding(retVal.IDispatch(), false, true)
}

