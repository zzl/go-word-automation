package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020982-0000-0000-C000-000000000046
var IID_Task = syscall.GUID{0x00020982, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Task struct {
	ole.OleClient
}

func NewTask(pDisp *win32.IDispatch, addRef bool, scoped bool) *Task {
	 if pDisp == nil {
		return nil;
	}
	p := &Task{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TaskFromVar(v ole.Variant) *Task {
	return NewTask(v.IDispatch(), false, false)
}

func (this *Task) IID() *syscall.GUID {
	return &IID_Task
}

func (this *Task) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Task) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Task) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Task) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Task) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Task) Left() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Task) SetLeft(rhs int32)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *Task) Top() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Task) SetTop(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *Task) Width() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Task) SetWidth(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Task) Height() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Task) SetHeight(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *Task) WindowState() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Task) SetWindowState(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *Task) Visible() bool {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Task) SetVisible(rhs bool)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

var Task_Activate_OptArgs= []string{
	"Wait", 
}

func (this *Task) Activate(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Task_Activate_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000a, nil, optArgs...)
	_= retVal
}

func (this *Task) Close()  {
	retVal, _ := this.Call(0x0000000b, nil)
	_= retVal
}

func (this *Task) Move(left int32, top int32)  {
	retVal, _ := this.Call(0x0000000c, []interface{}{left, top})
	_= retVal
}

func (this *Task) Resize(width int32, height int32)  {
	retVal, _ := this.Call(0x0000000d, []interface{}{width, height})
	_= retVal
}

func (this *Task) SendWindowMessage(message int32, wParam int32, lParam int32)  {
	retVal, _ := this.Call(0x0000000e, []interface{}{message, wParam, lParam})
	_= retVal
}

