package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020928-0000-0000-C000-000000000046
var IID_FormField = syscall.GUID{0x00020928, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FormField struct {
	ole.OleClient
}

func NewFormField(pDisp *win32.IDispatch, addRef bool, scoped bool) *FormField {
	p := &FormField{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FormFieldFromVar(v ole.Variant) *FormField {
	return NewFormField(v.PdispValVal(), false, false)
}

func (this *FormField) IID() *syscall.GUID {
	return &IID_FormField
}

func (this *FormField) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FormField) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *FormField) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *FormField) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormField) Type() int32 {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *FormField) Name() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FormField) SetName(rhs string)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *FormField) EntryMacro() string {
	retVal := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FormField) SetEntryMacro(rhs string)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *FormField) ExitMacro() string {
	retVal := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FormField) SetExitMacro(rhs string)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *FormField) OwnHelp() bool {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *FormField) SetOwnHelp(rhs bool)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *FormField) OwnStatus() bool {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *FormField) SetOwnStatus(rhs bool)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *FormField) HelpText() string {
	retVal := this.PropGet(0x00000007, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FormField) SetHelpText(rhs string)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *FormField) StatusText() string {
	retVal := this.PropGet(0x00000008, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FormField) SetStatusText(rhs string)  {
	retVal := this.PropPut(0x00000008, []interface{}{rhs})
	_= retVal
}

func (this *FormField) Enabled() bool {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *FormField) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *FormField) Result() string {
	retVal := this.PropGet(0x0000000a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FormField) SetResult(rhs string)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *FormField) TextInput() *TextInput {
	retVal := this.PropGet(0x0000000b, nil)
	return NewTextInput(retVal.PdispValVal(), false, true)
}

func (this *FormField) CheckBox() *CheckBox {
	retVal := this.PropGet(0x0000000c, nil)
	return NewCheckBox(retVal.PdispValVal(), false, true)
}

func (this *FormField) DropDown() *DropDown {
	retVal := this.PropGet(0x0000000d, nil)
	return NewDropDown(retVal.PdispValVal(), false, true)
}

func (this *FormField) Next() *FormField {
	retVal := this.PropGet(0x0000000e, nil)
	return NewFormField(retVal.PdispValVal(), false, true)
}

func (this *FormField) Previous() *FormField {
	retVal := this.PropGet(0x0000000f, nil)
	return NewFormField(retVal.PdispValVal(), false, true)
}

func (this *FormField) CalculateOnExit() bool {
	retVal := this.PropGet(0x00000010, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *FormField) SetCalculateOnExit(rhs bool)  {
	retVal := this.PropPut(0x00000010, []interface{}{rhs})
	_= retVal
}

func (this *FormField) Range() *Range {
	retVal := this.PropGet(0x00000011, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *FormField) Select()  {
	retVal := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *FormField) Copy()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *FormField) Cut()  {
	retVal := this.Call(0x00000066, nil)
	_= retVal
}

func (this *FormField) Delete()  {
	retVal := this.Call(0x00000067, nil)
	_= retVal
}

