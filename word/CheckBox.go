package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020926-0000-0000-C000-000000000046
var IID_CheckBox = syscall.GUID{0x00020926, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CheckBox struct {
	ole.OleClient
}

func NewCheckBox(pDisp *win32.IDispatch, addRef bool, scoped bool) *CheckBox {
	p := &CheckBox{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CheckBoxFromVar(v ole.Variant) *CheckBox {
	return NewCheckBox(v.PdispValVal(), false, false)
}

func (this *CheckBox) IID() *syscall.GUID {
	return &IID_CheckBox
}

func (this *CheckBox) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CheckBox) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CheckBox) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CheckBox) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CheckBox) Valid() bool {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CheckBox) AutoSize() bool {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CheckBox) SetAutoSize(rhs bool)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *CheckBox) Size() float32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.FltValVal()
}

func (this *CheckBox) SetSize(rhs float32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *CheckBox) Default() bool {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CheckBox) SetDefault(rhs bool)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *CheckBox) Value() bool {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CheckBox) SetValue(rhs bool)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

