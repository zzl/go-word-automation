package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002097E-0000-0000-C000-000000000046
var IID_AddIn = syscall.GUID{0x0002097E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AddIn struct {
	ole.OleClient
}

func NewAddIn(pDisp *win32.IDispatch, addRef bool, scoped bool) *AddIn {
	 if pDisp == nil {
		return nil;
	}
	p := &AddIn{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AddInFromVar(v ole.Variant) *AddIn {
	return NewAddIn(v.IDispatch(), false, false)
}

func (this *AddIn) IID() *syscall.GUID {
	return &IID_AddIn
}

func (this *AddIn) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AddIn) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AddIn) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AddIn) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AddIn) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *AddIn) Path() string {
	retVal, _ := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AddIn) Installed() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AddIn) SetInstalled(rhs bool)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *AddIn) Compiled() bool {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AddIn) Autoload() bool {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AddIn) Delete()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

