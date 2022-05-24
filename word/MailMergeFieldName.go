package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002091B-0000-0000-C000-000000000046
var IID_MailMergeFieldName = syscall.GUID{0x0002091B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailMergeFieldName struct {
	ole.OleClient
}

func NewMailMergeFieldName(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailMergeFieldName {
	 if pDisp == nil {
		return nil;
	}
	p := &MailMergeFieldName{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailMergeFieldNameFromVar(v ole.Variant) *MailMergeFieldName {
	return NewMailMergeFieldName(v.IDispatch(), false, false)
}

func (this *MailMergeFieldName) IID() *syscall.GUID {
	return &IID_MailMergeFieldName
}

func (this *MailMergeFieldName) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailMergeFieldName) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *MailMergeFieldName) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailMergeFieldName) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *MailMergeFieldName) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMergeFieldName) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

