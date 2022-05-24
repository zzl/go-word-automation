package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020945-0000-0000-C000-000000000046
var IID_FirstLetterException = syscall.GUID{0x00020945, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FirstLetterException struct {
	ole.OleClient
}

func NewFirstLetterException(pDisp *win32.IDispatch, addRef bool, scoped bool) *FirstLetterException {
	 if pDisp == nil {
		return nil;
	}
	p := &FirstLetterException{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FirstLetterExceptionFromVar(v ole.Variant) *FirstLetterException {
	return NewFirstLetterException(v.IDispatch(), false, false)
}

func (this *FirstLetterException) IID() *syscall.GUID {
	return &IID_FirstLetterException
}

func (this *FirstLetterException) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FirstLetterException) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *FirstLetterException) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *FirstLetterException) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *FirstLetterException) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *FirstLetterException) Name() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *FirstLetterException) Delete()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

