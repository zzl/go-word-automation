package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209D2-0000-0000-C000-000000000046
var IID_HangulAndAlphabetException = syscall.GUID{0x000209D2, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HangulAndAlphabetException struct {
	ole.OleClient
}

func NewHangulAndAlphabetException(pDisp *win32.IDispatch, addRef bool, scoped bool) *HangulAndAlphabetException {
	p := &HangulAndAlphabetException{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HangulAndAlphabetExceptionFromVar(v ole.Variant) *HangulAndAlphabetException {
	return NewHangulAndAlphabetException(v.PdispValVal(), false, false)
}

func (this *HangulAndAlphabetException) IID() *syscall.GUID {
	return &IID_HangulAndAlphabetException
}

func (this *HangulAndAlphabetException) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HangulAndAlphabetException) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *HangulAndAlphabetException) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HangulAndAlphabetException) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *HangulAndAlphabetException) Index() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *HangulAndAlphabetException) Name() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *HangulAndAlphabetException) Delete()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

