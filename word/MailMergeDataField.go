package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020919-0000-0000-C000-000000000046
var IID_MailMergeDataField = syscall.GUID{0x00020919, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailMergeDataField struct {
	ole.OleClient
}

func NewMailMergeDataField(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailMergeDataField {
	p := &MailMergeDataField{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailMergeDataFieldFromVar(v ole.Variant) *MailMergeDataField {
	return NewMailMergeDataField(v.PdispValVal(), false, false)
}

func (this *MailMergeDataField) IID() *syscall.GUID {
	return &IID_MailMergeDataField
}

func (this *MailMergeDataField) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailMergeDataField) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MailMergeDataField) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataField) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MailMergeDataField) Value() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMergeDataField) Name() string {
	retVal := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMergeDataField) Index() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

