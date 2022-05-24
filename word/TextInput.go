package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020927-0000-0000-C000-000000000046
var IID_TextInput = syscall.GUID{0x00020927, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TextInput struct {
	ole.OleClient
}

func NewTextInput(pDisp *win32.IDispatch, addRef bool, scoped bool) *TextInput {
	 if pDisp == nil {
		return nil;
	}
	p := &TextInput{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TextInputFromVar(v ole.Variant) *TextInput {
	return NewTextInput(v.IDispatch(), false, false)
}

func (this *TextInput) IID() *syscall.GUID {
	return &IID_TextInput
}

func (this *TextInput) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TextInput) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TextInput) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TextInput) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextInput) Valid() bool {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextInput) Default() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TextInput) SetDefault(rhs string)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *TextInput) Type() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *TextInput) Format() string {
	retVal, _ := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TextInput) Width() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *TextInput) SetWidth(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *TextInput) Clear()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

var TextInput_EditType_OptArgs= []string{
	"Default", "Format", "Enabled", 
}

func (this *TextInput) EditType(type_ int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(TextInput_EditType_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, []interface{}{type_}, optArgs...)
	_= retVal
}

