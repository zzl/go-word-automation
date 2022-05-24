package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020998-0000-0000-C000-000000000046
var IID_KeyBinding = syscall.GUID{0x00020998, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type KeyBinding struct {
	ole.OleClient
}

func NewKeyBinding(pDisp *win32.IDispatch, addRef bool, scoped bool) *KeyBinding {
	 if pDisp == nil {
		return nil;
	}
	p := &KeyBinding{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func KeyBindingFromVar(v ole.Variant) *KeyBinding {
	return NewKeyBinding(v.IDispatch(), false, false)
}

func (this *KeyBinding) IID() *syscall.GUID {
	return &IID_KeyBinding
}

func (this *KeyBinding) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *KeyBinding) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *KeyBinding) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *KeyBinding) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *KeyBinding) Command() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *KeyBinding) KeyString() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *KeyBinding) Protected() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *KeyBinding) KeyCategory() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *KeyBinding) KeyCode() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *KeyBinding) KeyCode2() int32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *KeyBinding) CommandParameter() string {
	retVal, _ := this.PropGet(0x00000008, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *KeyBinding) Context() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *KeyBinding) Clear()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *KeyBinding) Disable()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *KeyBinding) Execute()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

var KeyBinding_Rebind_OptArgs= []string{
	"CommandParameter", 
}

func (this *KeyBinding) Rebind(keyCategory int32, command string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(KeyBinding_Rebind_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000068, []interface{}{keyCategory, command}, optArgs...)
	_= retVal
}

