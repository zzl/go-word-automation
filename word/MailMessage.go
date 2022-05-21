package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209BA-0000-0000-C000-000000000046
var IID_MailMessage = syscall.GUID{0x000209BA, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailMessage struct {
	ole.OleClient
}

func NewMailMessage(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailMessage {
	p := &MailMessage{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailMessageFromVar(v ole.Variant) *MailMessage {
	return NewMailMessage(v.PdispValVal(), false, false)
}

func (this *MailMessage) IID() *syscall.GUID {
	return &IID_MailMessage
}

func (this *MailMessage) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailMessage) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MailMessage) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailMessage) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MailMessage) CheckName()  {
	retVal := this.Call(0x0000014e, nil)
	_= retVal
}

func (this *MailMessage) Delete()  {
	retVal := this.Call(0x0000014f, nil)
	_= retVal
}

func (this *MailMessage) DisplayMoveDialog()  {
	retVal := this.Call(0x00000150, nil)
	_= retVal
}

func (this *MailMessage) DisplayProperties()  {
	retVal := this.Call(0x00000151, nil)
	_= retVal
}

func (this *MailMessage) DisplaySelectNamesDialog()  {
	retVal := this.Call(0x00000152, nil)
	_= retVal
}

func (this *MailMessage) Forward()  {
	retVal := this.Call(0x00000153, nil)
	_= retVal
}

func (this *MailMessage) GoToNext()  {
	retVal := this.Call(0x00000154, nil)
	_= retVal
}

func (this *MailMessage) GoToPrevious()  {
	retVal := this.Call(0x00000155, nil)
	_= retVal
}

func (this *MailMessage) Reply()  {
	retVal := this.Call(0x00000156, nil)
	_= retVal
}

func (this *MailMessage) ReplyAll()  {
	retVal := this.Call(0x00000157, nil)
	_= retVal
}

func (this *MailMessage) ToggleHeader()  {
	retVal := this.Call(0x00000158, nil)
	_= retVal
}

