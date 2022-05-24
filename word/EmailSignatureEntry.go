package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209E6-0000-0000-C000-000000000046
var IID_EmailSignatureEntry = syscall.GUID{0x000209E6, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type EmailSignatureEntry struct {
	ole.OleClient
}

func NewEmailSignatureEntry(pDisp *win32.IDispatch, addRef bool, scoped bool) *EmailSignatureEntry {
	 if pDisp == nil {
		return nil;
	}
	p := &EmailSignatureEntry{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EmailSignatureEntryFromVar(v ole.Variant) *EmailSignatureEntry {
	return NewEmailSignatureEntry(v.IDispatch(), false, false)
}

func (this *EmailSignatureEntry) IID() *syscall.GUID {
	return &IID_EmailSignatureEntry
}

func (this *EmailSignatureEntry) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *EmailSignatureEntry) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *EmailSignatureEntry) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *EmailSignatureEntry) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *EmailSignatureEntry) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *EmailSignatureEntry) Name() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EmailSignatureEntry) SetName(rhs string)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *EmailSignatureEntry) Delete()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

