package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209B8-0000-0000-C000-000000000046
var IID_Dialog = syscall.GUID{0x000209B8, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Dialog struct {
	ole.OleClient
}

func NewDialog(pDisp *win32.IDispatch, addRef bool, scoped bool) *Dialog {
	p := &Dialog{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DialogFromVar(v ole.Variant) *Dialog {
	return NewDialog(v.PdispValVal(), false, false)
}

func (this *Dialog) IID() *syscall.GUID {
	return &IID_Dialog
}

func (this *Dialog) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Dialog) Application() *Application {
	retVal := this.PropGet(0x00007d03, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Dialog) Creator() int32 {
	retVal := this.PropGet(0x00007d04, nil)
	return retVal.LValVal()
}

func (this *Dialog) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00007d05, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Dialog) DefaultTab() int32 {
	retVal := this.PropGet(0x00007d02, nil)
	return retVal.LValVal()
}

func (this *Dialog) SetDefaultTab(rhs int32)  {
	retVal := this.PropPut(0x00007d02, []interface{}{rhs})
	_= retVal
}

func (this *Dialog) Type() int32 {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

var Dialog_Show_OptArgs= []string{
	"TimeOut", 
}

func (this *Dialog) Show(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Dialog_Show_OptArgs, optArgs)
	retVal := this.Call(0x00000150, nil, optArgs...)
	return retVal.LValVal()
}

var Dialog_Display_OptArgs= []string{
	"TimeOut", 
}

func (this *Dialog) Display(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Dialog_Display_OptArgs, optArgs)
	retVal := this.Call(0x00000152, nil, optArgs...)
	return retVal.LValVal()
}

func (this *Dialog) Execute()  {
	retVal := this.Call(0x00007d01, nil)
	_= retVal
}

func (this *Dialog) Update()  {
	retVal := this.Call(0x0000012e, nil)
	_= retVal
}

func (this *Dialog) CommandName() string {
	retVal := this.PropGet(0x00007d06, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Dialog) CommandBarId() int32 {
	retVal := this.PropGet(0x00007d07, nil)
	return retVal.LValVal()
}

