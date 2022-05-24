package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020A00-0001-0000-C000-000000000046
var IID_IApplicationEvents3 = syscall.GUID{0x00020A00, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IApplicationEvents3 struct {
	ole.OleClient
}

func NewIApplicationEvents3(pDisp *win32.IDispatch, addRef bool, scoped bool) *IApplicationEvents3 {
	 if pDisp == nil {
		return nil;
	}
	p := &IApplicationEvents3{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func IApplicationEvents3FromVar(v ole.Variant) *IApplicationEvents3 {
	return NewIApplicationEvents3(v.IDispatch(), false, false)
}

func (this *IApplicationEvents3) IID() *syscall.GUID {
	return &IID_IApplicationEvents3
}

func (this *IApplicationEvents3) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *IApplicationEvents3) Startup()  {
	retVal, _ := this.Call(0x00000001, nil)
	_= retVal
}

func (this *IApplicationEvents3) Quit()  {
	retVal, _ := this.Call(0x00000002, nil)
	_= retVal
}

func (this *IApplicationEvents3) DocumentChange()  {
	retVal, _ := this.Call(0x00000003, nil)
	_= retVal
}

func (this *IApplicationEvents3) DocumentOpen(doc *Document)  {
	retVal, _ := this.Call(0x00000004, []interface{}{doc})
	_= retVal
}

func (this *IApplicationEvents3) DocumentBeforeClose(doc *Document, cancel *win32.VARIANT_BOOL)  {
	retVal, _ := this.Call(0x00000006, []interface{}{doc, cancel})
	_= retVal
}

func (this *IApplicationEvents3) DocumentBeforePrint(doc *Document, cancel *win32.VARIANT_BOOL)  {
	retVal, _ := this.Call(0x00000007, []interface{}{doc, cancel})
	_= retVal
}

func (this *IApplicationEvents3) DocumentBeforeSave(doc *Document, saveAsUI *win32.VARIANT_BOOL, cancel *win32.VARIANT_BOOL)  {
	retVal, _ := this.Call(0x00000008, []interface{}{doc, saveAsUI, cancel})
	_= retVal
}

func (this *IApplicationEvents3) NewDocument(doc *Document)  {
	retVal, _ := this.Call(0x00000009, []interface{}{doc})
	_= retVal
}

func (this *IApplicationEvents3) WindowActivate(doc *Document, wn *Window)  {
	retVal, _ := this.Call(0x0000000a, []interface{}{doc, wn})
	_= retVal
}

func (this *IApplicationEvents3) WindowDeactivate(doc *Document, wn *Window)  {
	retVal, _ := this.Call(0x0000000b, []interface{}{doc, wn})
	_= retVal
}

func (this *IApplicationEvents3) WindowSelectionChange(sel *Selection)  {
	retVal, _ := this.Call(0x0000000c, []interface{}{sel})
	_= retVal
}

func (this *IApplicationEvents3) WindowBeforeRightClick(sel *Selection, cancel *win32.VARIANT_BOOL)  {
	retVal, _ := this.Call(0x0000000d, []interface{}{sel, cancel})
	_= retVal
}

func (this *IApplicationEvents3) WindowBeforeDoubleClick(sel *Selection, cancel *win32.VARIANT_BOOL)  {
	retVal, _ := this.Call(0x0000000e, []interface{}{sel, cancel})
	_= retVal
}

func (this *IApplicationEvents3) EPostagePropertyDialog(doc *Document)  {
	retVal, _ := this.Call(0x0000000f, []interface{}{doc})
	_= retVal
}

func (this *IApplicationEvents3) EPostageInsert(doc *Document)  {
	retVal, _ := this.Call(0x00000010, []interface{}{doc})
	_= retVal
}

func (this *IApplicationEvents3) MailMergeAfterMerge(doc *Document, docResult *Document)  {
	retVal, _ := this.Call(0x00000011, []interface{}{doc, docResult})
	_= retVal
}

func (this *IApplicationEvents3) MailMergeAfterRecordMerge(doc *Document)  {
	retVal, _ := this.Call(0x00000012, []interface{}{doc})
	_= retVal
}

func (this *IApplicationEvents3) MailMergeBeforeMerge(doc *Document, startRecord int32, endRecord int32, cancel *win32.VARIANT_BOOL)  {
	retVal, _ := this.Call(0x00000013, []interface{}{doc, startRecord, endRecord, cancel})
	_= retVal
}

func (this *IApplicationEvents3) MailMergeBeforeRecordMerge(doc *Document, cancel *win32.VARIANT_BOOL)  {
	retVal, _ := this.Call(0x00000014, []interface{}{doc, cancel})
	_= retVal
}

func (this *IApplicationEvents3) MailMergeDataSourceLoad(doc *Document)  {
	retVal, _ := this.Call(0x00000015, []interface{}{doc})
	_= retVal
}

func (this *IApplicationEvents3) MailMergeDataSourceValidate(doc *Document, handled *win32.VARIANT_BOOL)  {
	retVal, _ := this.Call(0x00000016, []interface{}{doc, handled})
	_= retVal
}

func (this *IApplicationEvents3) MailMergeWizardSendToCustom(doc *Document)  {
	retVal, _ := this.Call(0x00000017, []interface{}{doc})
	_= retVal
}

func (this *IApplicationEvents3) MailMergeWizardStateChange(doc *Document, fromState *int32, toState *int32, handled *win32.VARIANT_BOOL)  {
	retVal, _ := this.Call(0x00000018, []interface{}{doc, fromState, toState, handled})
	_= retVal
}

func (this *IApplicationEvents3) WindowSize(doc *Document, wn *Window)  {
	retVal, _ := this.Call(0x00000019, []interface{}{doc, wn})
	_= retVal
}

