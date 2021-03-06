package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209DC-0000-0000-C000-000000000046
var IID_EmailSignature = syscall.GUID{0x000209DC, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type EmailSignature struct {
	ole.OleClient
}

func NewEmailSignature(pDisp *win32.IDispatch, addRef bool, scoped bool) *EmailSignature {
	 if pDisp == nil {
		return nil;
	}
	p := &EmailSignature{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EmailSignatureFromVar(v ole.Variant) *EmailSignature {
	return NewEmailSignature(v.IDispatch(), false, false)
}

func (this *EmailSignature) IID() *syscall.GUID {
	return &IID_EmailSignature
}

func (this *EmailSignature) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *EmailSignature) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *EmailSignature) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *EmailSignature) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *EmailSignature) NewMessageSignature() string {
	retVal, _ := this.PropGet(0x00000067, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EmailSignature) SetNewMessageSignature(rhs string)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *EmailSignature) ReplyMessageSignature() string {
	retVal, _ := this.PropGet(0x00000068, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EmailSignature) SetReplyMessageSignature(rhs string)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *EmailSignature) EmailSignatureEntries() *EmailSignatureEntries {
	retVal, _ := this.PropGet(0x00000069, nil)
	return NewEmailSignatureEntries(retVal.IDispatch(), false, true)
}

