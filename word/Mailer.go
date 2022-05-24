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
	 if pDisp == nil {
		return nil;
	}
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
	return NewMailer(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Mailer) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Mailer) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Mailer) BCCRecipients() ole.Variant {
	retVal, _ := this.PropGet(0x00000064, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Mailer) SetBCCRecipients(rhs *ole.Variant)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *Mailer) CCRecipients() ole.Variant {
	retVal, _ := this.PropGet(0x00000065, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Mailer) SetCCRecipients(rhs *ole.Variant)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *Mailer) Recipients() ole.Variant {
	retVal, _ := this.PropGet(0x00000066, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Mailer) SetRecipients(rhs *ole.Variant)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *Mailer) Enclosures() ole.Variant {
	retVal, _ := this.PropGet(0x00000067, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Mailer) SetEnclosures(rhs *ole.Variant)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Mailer) Sender() string {
	retVal, _ := this.PropGet(0x00000068, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Mailer) SendDateTime() time.Time {
	retVal, _ := this.PropGet(0x00000069, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *Mailer) Received() bool {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Mailer) Subject() string {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Mailer) SetSubject(rhs string)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

