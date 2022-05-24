package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020943-0000-0000-C000-000000000046
var IID_TwoInitialCapsException = syscall.GUID{0x00020943, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TwoInitialCapsException struct {
	ole.OleClient
}

func NewTwoInitialCapsException(pDisp *win32.IDispatch, addRef bool, scoped bool) *TwoInitialCapsException {
	 if pDisp == nil {
		return nil;
	}
	p := &TwoInitialCapsException{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TwoInitialCapsExceptionFromVar(v ole.Variant) *TwoInitialCapsException {
	return NewTwoInitialCapsException(v.IDispatch(), false, false)
}

func (this *TwoInitialCapsException) IID() *syscall.GUID {
	return &IID_TwoInitialCapsException
}

func (this *TwoInitialCapsException) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TwoInitialCapsException) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TwoInitialCapsException) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TwoInitialCapsException) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TwoInitialCapsException) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *TwoInitialCapsException) Name() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TwoInitialCapsException) Delete()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

