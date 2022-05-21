package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209D7-0000-0000-C000-000000000046
var IID_EmailAuthor = syscall.GUID{0x000209D7, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type EmailAuthor struct {
	ole.OleClient
}

func NewEmailAuthor(pDisp *win32.IDispatch, addRef bool, scoped bool) *EmailAuthor {
	p := &EmailAuthor{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EmailAuthorFromVar(v ole.Variant) *EmailAuthor {
	return NewEmailAuthor(v.PdispValVal(), false, false)
}

func (this *EmailAuthor) IID() *syscall.GUID {
	return &IID_EmailAuthor
}

func (this *EmailAuthor) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *EmailAuthor) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *EmailAuthor) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *EmailAuthor) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *EmailAuthor) Style() *Style {
	retVal := this.PropGet(0x00000067, nil)
	return NewStyle(retVal.PdispValVal(), false, true)
}

