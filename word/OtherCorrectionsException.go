package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209E1-0000-0000-C000-000000000046
var IID_OtherCorrectionsException = syscall.GUID{0x000209E1, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OtherCorrectionsException struct {
	ole.OleClient
}

func NewOtherCorrectionsException(pDisp *win32.IDispatch, addRef bool, scoped bool) *OtherCorrectionsException {
	p := &OtherCorrectionsException{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OtherCorrectionsExceptionFromVar(v ole.Variant) *OtherCorrectionsException {
	return NewOtherCorrectionsException(v.PdispValVal(), false, false)
}

func (this *OtherCorrectionsException) IID() *syscall.GUID {
	return &IID_OtherCorrectionsException
}

func (this *OtherCorrectionsException) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OtherCorrectionsException) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OtherCorrectionsException) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *OtherCorrectionsException) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OtherCorrectionsException) Index() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *OtherCorrectionsException) Name() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OtherCorrectionsException) Delete()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

