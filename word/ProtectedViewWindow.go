package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// F743EDD0-9B97-4B09-89CC-77BE19B51481
var IID_ProtectedViewWindow = syscall.GUID{0xF743EDD0, 0x9B97, 0x4B09, 
	[8]byte{0x89, 0xCC, 0x77, 0xBE, 0x19, 0xB5, 0x14, 0x81}}

type ProtectedViewWindow struct {
	ole.OleClient
}

func NewProtectedViewWindow(pDisp *win32.IDispatch, addRef bool, scoped bool) *ProtectedViewWindow {
	 if pDisp == nil {
		return nil;
	}
	p := &ProtectedViewWindow{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ProtectedViewWindowFromVar(v ole.Variant) *ProtectedViewWindow {
	return NewProtectedViewWindow(v.IDispatch(), false, false)
}

func (this *ProtectedViewWindow) IID() *syscall.GUID {
	return &IID_ProtectedViewWindow
}

func (this *ProtectedViewWindow) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ProtectedViewWindow) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ProtectedViewWindow) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ProtectedViewWindow) Caption() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) SetCaption(rhs string)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *ProtectedViewWindow) Document() *Document {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewDocument(retVal.IDispatch(), false, true)
}

func (this *ProtectedViewWindow) Left() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetLeft(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *ProtectedViewWindow) Top() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetTop(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *ProtectedViewWindow) Width() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetWidth(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *ProtectedViewWindow) Height() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetHeight(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *ProtectedViewWindow) WindowState() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetWindowState(rhs int32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *ProtectedViewWindow) Active() bool {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ProtectedViewWindow) Index() int32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) Visible() bool {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ProtectedViewWindow) SetVisible(rhs bool)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *ProtectedViewWindow) SourceName() string {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) SourcePath() string {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) Activate()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

var ProtectedViewWindow_Edit_OptArgs= []string{
	"PasswordTemplate", "WritePasswordDocument", "WritePasswordTemplate", 
}

func (this *ProtectedViewWindow) Edit(optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(ProtectedViewWindow_Edit_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, nil, optArgs...)
	return NewDocument(retVal.IDispatch(), false, true)
}

func (this *ProtectedViewWindow) Close()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *ProtectedViewWindow) ToggleRibbon()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

