package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209EF-0000-0000-C000-000000000046
var IID_StyleSheet = syscall.GUID{0x000209EF, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type StyleSheet struct {
	ole.OleClient
}

func NewStyleSheet(pDisp *win32.IDispatch, addRef bool, scoped bool) *StyleSheet {
	 if pDisp == nil {
		return nil;
	}
	p := &StyleSheet{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func StyleSheetFromVar(v ole.Variant) *StyleSheet {
	return NewStyleSheet(v.IDispatch(), false, false)
}

func (this *StyleSheet) IID() *syscall.GUID {
	return &IID_StyleSheet
}

func (this *StyleSheet) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *StyleSheet) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *StyleSheet) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *StyleSheet) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000064, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *StyleSheet) FullName() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *StyleSheet) Index() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *StyleSheet) Name() string {
	retVal, _ := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *StyleSheet) Path() string {
	retVal, _ := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *StyleSheet) Type() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *StyleSheet) SetType(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *StyleSheet) Title() string {
	retVal, _ := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *StyleSheet) SetTitle(rhs string)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *StyleSheet) Move(precedence int32)  {
	retVal, _ := this.Call(0x00000007, []interface{}{precedence})
	_= retVal
}

func (this *StyleSheet) Delete()  {
	retVal, _ := this.Call(0x00000008, nil)
	_= retVal
}

