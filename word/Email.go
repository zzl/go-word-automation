package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209DD-0000-0000-C000-000000000046
var IID_Email = syscall.GUID{0x000209DD, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Email struct {
	ole.OleClient
}

func NewEmail(pDisp *win32.IDispatch, addRef bool, scoped bool) *Email {
	p := &Email{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EmailFromVar(v ole.Variant) *Email {
	return NewEmail(v.PdispValVal(), false, false)
}

func (this *Email) IID() *syscall.GUID {
	return &IID_Email
}

func (this *Email) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Email) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Email) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Email) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Email) CurrentEmailAuthor() *EmailAuthor {
	retVal := this.PropGet(0x00000069, nil)
	return NewEmailAuthor(retVal.PdispValVal(), false, true)
}

