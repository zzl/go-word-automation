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
	return NewProtectedViewWindow(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ProtectedViewWindow) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ProtectedViewWindow) Caption() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) SetCaption(rhs string)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Document() *Document {
	retVal := this.PropGet(0x00000001, nil)
	return NewDocument(retVal.PdispValVal(), false, true)
}

func (this *ProtectedViewWindow) Left() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetLeft(rhs int32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Top() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetTop(rhs int32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Width() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetWidth(rhs int32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Height() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetHeight(rhs int32)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) WindowState() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) SetWindowState(rhs int32)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) Active() bool {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ProtectedViewWindow) Index() int32 {
	retVal := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindow) Visible() bool {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ProtectedViewWindow) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *ProtectedViewWindow) SourceName() string {
	retVal := this.PropGet(0x0000000a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) SourcePath() string {
	retVal := this.PropGet(0x0000000b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ProtectedViewWindow) Activate()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

var ProtectedViewWindow_Edit_OptArgs= []string{
	"PasswordTemplate", "WritePasswordDocument", "WritePasswordTemplate", 
}

func (this *ProtectedViewWindow) Edit(optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(ProtectedViewWindow_Edit_OptArgs, optArgs)
	retVal := this.Call(0x00000065, nil, optArgs...)
	return NewDocument(retVal.PdispValVal(), false, true)
}

func (this *ProtectedViewWindow) Close()  {
	retVal := this.Call(0x00000066, nil)
	_= retVal
}

func (this *ProtectedViewWindow) ToggleRibbon()  {
	retVal := this.Call(0x00000067, nil)
	_= retVal
}

