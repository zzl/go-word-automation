package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"time"
)

// 000209BD-0000-0000-C000-000000000046
var IID_Mailer = syscall.GUID{0x000209BD, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Mailer struct {
	ole.OleClient
}

func NewMailer(pDisp *win32.IDispatch, addRef bool, scoped bool) *Mailer {
	p := &Mailer{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailerFromVar(v ole.Variant) *Mailer {
	return NewMailer(v.PdispValVal(), false, false)
}

func (this *Mailer) IID() *syscall.GUID {
	return &IID_Mailer
}

func (this *Mailer) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Mailer) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Mailer) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Mailer) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Mailer) BCCRecipients() ole.Variant {
	retVal := this.PropGet(0x00000064, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Mailer) SetBCCRecipients(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *Mailer) CCRecipients() ole.Variant {
	retVal := this.PropGet(0x00000065, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Mailer) SetCCRecipients(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *Mailer) Recipients() ole.Variant {
	retVal := this.PropGet(0x00000066, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Mailer) SetRecipients(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *Mailer) Enclosures() ole.Variant {
	retVal := this.PropGet(0x00000067, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Mailer) SetEnclosures(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *Mailer) Sender() string {
	retVal := this.PropGet(0x00000068, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Mailer) SendDateTime() time.Time {
	retVal := this.PropGet(0x00000069, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *Mailer) Received() bool {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Mailer) Subject() string {
	retVal := this.PropGet(0x0000006b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Mailer) SetSubject(rhs string)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

