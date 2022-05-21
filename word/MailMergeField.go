package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002091E-0000-0000-C000-000000000046
var IID_MailMergeField = syscall.GUID{0x0002091E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailMergeField struct {
	ole.OleClient
}

func NewMailMergeField(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailMergeField {
	p := &MailMergeField{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailMergeFieldFromVar(v ole.Variant) *MailMergeField {
	return NewMailMergeField(v.PdispValVal(), false, false)
}

func (this *MailMergeField) IID() *syscall.GUID {
	return &IID_MailMergeField
}

func (this *MailMergeField) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailMergeField) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MailMergeField) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailMergeField) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MailMergeField) Type() int32 {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *MailMergeField) Locked() bool {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailMergeField) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *MailMergeField) Code() *Range {
	retVal := this.PropGet(0x00000005, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *MailMergeField) SetCode(rhs *Range)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *MailMergeField) Next() *MailMergeField {
	retVal := this.PropGet(0x00000008, nil)
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

func (this *MailMergeField) Previous() *MailMergeField {
	retVal := this.PropGet(0x00000009, nil)
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

func (this *MailMergeField) Select()  {
	retVal := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *MailMergeField) Copy()  {
	retVal := this.Call(0x00000069, nil)
	_= retVal
}

func (this *MailMergeField) Cut()  {
	retVal := this.Call(0x0000006a, nil)
	_= retVal
}

func (this *MailMergeField) Delete()  {
	retVal := this.Call(0x0000006b, nil)
	_= retVal
}

