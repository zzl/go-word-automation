package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020A00-0000-0000-C000-000000000046
var IID_ApplicationEvents3 = syscall.GUID{0x00020A00, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ApplicationEvents3DispInterface interface {
	QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) 
	AddRef_() uint32
	Release_() uint32
	GetTypeInfoCount_(pctinfo *uint32) 
	GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) 
	GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) 
	Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) 
	Startup() 
	Quit() 
	DocumentChange() 
	DocumentOpen(doc *Document) 
	DocumentBeforeClose(doc *Document, cancel *win32.VARIANT_BOOL) 
	DocumentBeforePrint(doc *Document, cancel *win32.VARIANT_BOOL) 
	DocumentBeforeSave(doc *Document, saveAsUI *win32.VARIANT_BOOL, cancel *win32.VARIANT_BOOL) 
	NewDocument(doc *Document) 
	WindowActivate(doc *Document, wn *Window) 
	WindowDeactivate(doc *Document, wn *Window) 
	WindowSelectionChange(sel *Selection) 
	WindowBeforeRightClick(sel *Selection, cancel *win32.VARIANT_BOOL) 
	WindowBeforeDoubleClick(sel *Selection, cancel *win32.VARIANT_BOOL) 
	EPostagePropertyDialog(doc *Document) 
	EPostageInsert(doc *Document) 
	MailMergeAfterMerge(doc *Document, docResult *Document) 
	MailMergeAfterRecordMerge(doc *Document) 
	MailMergeBeforeMerge(doc *Document, startRecord int32, endRecord int32, cancel *win32.VARIANT_BOOL) 
	MailMergeBeforeRecordMerge(doc *Document, cancel *win32.VARIANT_BOOL) 
	MailMergeDataSourceLoad(doc *Document) 
	MailMergeDataSourceValidate(doc *Document, handled *win32.VARIANT_BOOL) 
	MailMergeWizardSendToCustom(doc *Document) 
	MailMergeWizardStateChange(doc *Document, fromState *int32, toState *int32, handled *win32.VARIANT_BOOL) 
	WindowSize(doc *Document, wn *Window) 
}

type ApplicationEvents3Handlers struct {
	QueryInterface_ func(riid *syscall.GUID, ppvObj unsafe.Pointer) 
	AddRef_ func() uint32
	Release_ func() uint32
	GetTypeInfoCount_ func(pctinfo *uint32) 
	GetTypeInfo_ func(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) 
	GetIDsOfNames_ func(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) 
	Invoke_ func(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) 
	Startup func() 
	Quit func() 
	DocumentChange func() 
	DocumentOpen func(doc *Document) 
	DocumentBeforeClose func(doc *Document, cancel *win32.VARIANT_BOOL) 
	DocumentBeforePrint func(doc *Document, cancel *win32.VARIANT_BOOL) 
	DocumentBeforeSave func(doc *Document, saveAsUI *win32.VARIANT_BOOL, cancel *win32.VARIANT_BOOL) 
	NewDocument func(doc *Document) 
	WindowActivate func(doc *Document, wn *Window) 
	WindowDeactivate func(doc *Document, wn *Window) 
	WindowSelectionChange func(sel *Selection) 
	WindowBeforeRightClick func(sel *Selection, cancel *win32.VARIANT_BOOL) 
	WindowBeforeDoubleClick func(sel *Selection, cancel *win32.VARIANT_BOOL) 
	EPostagePropertyDialog func(doc *Document) 
	EPostageInsert func(doc *Document) 
	MailMergeAfterMerge func(doc *Document, docResult *Document) 
	MailMergeAfterRecordMerge func(doc *Document) 
	MailMergeBeforeMerge func(doc *Document, startRecord int32, endRecord int32, cancel *win32.VARIANT_BOOL) 
	MailMergeBeforeRecordMerge func(doc *Document, cancel *win32.VARIANT_BOOL) 
	MailMergeDataSourceLoad func(doc *Document) 
	MailMergeDataSourceValidate func(doc *Document, handled *win32.VARIANT_BOOL) 
	MailMergeWizardSendToCustom func(doc *Document) 
	MailMergeWizardStateChange func(doc *Document, fromState *int32, toState *int32, handled *win32.VARIANT_BOOL) 
	WindowSize func(doc *Document, wn *Window) 
}

type ApplicationEvents3DispImpl struct {
	Handlers ApplicationEvents3Handlers
}

func (this *ApplicationEvents3DispImpl) QueryInterface_(riid *syscall.GUID, ppvObj unsafe.Pointer) {
	if this.Handlers.QueryInterface_ != nil {
		this.Handlers.QueryInterface_(riid, ppvObj)
	}
}

func (this *ApplicationEvents3DispImpl) AddRef_() uint32 {
	if this.Handlers.AddRef_ != nil {
		return this.Handlers.AddRef_()
	}
	var ret uint32
	return ret
}

func (this *ApplicationEvents3DispImpl) Release_() uint32 {
	if this.Handlers.Release_ != nil {
		return this.Handlers.Release_()
	}
	var ret uint32
	return ret
}

func (this *ApplicationEvents3DispImpl) GetTypeInfoCount_(pctinfo *uint32) {
	if this.Handlers.GetTypeInfoCount_ != nil {
		this.Handlers.GetTypeInfoCount_(pctinfo)
	}
}

func (this *ApplicationEvents3DispImpl) GetTypeInfo_(itinfo uint32, lcid uint32, pptinfo unsafe.Pointer) {
	if this.Handlers.GetTypeInfo_ != nil {
		this.Handlers.GetTypeInfo_(itinfo, lcid, pptinfo)
	}
}

func (this *ApplicationEvents3DispImpl) GetIDsOfNames_(riid *syscall.GUID, rgszNames **int8, cNames uint32, lcid uint32, rgdispid *int32) {
	if this.Handlers.GetIDsOfNames_ != nil {
		this.Handlers.GetIDsOfNames_(riid, rgszNames, cNames, lcid, rgdispid)
	}
}

func (this *ApplicationEvents3DispImpl) Invoke_(dispidMember int32, riid *syscall.GUID, lcid uint32, wFlags uint16, pdispparams *win32.DISPPARAMS, pvarResult *ole.Variant, pexcepinfo *win32.EXCEPINFO, puArgErr *uint32) {
	if this.Handlers.Invoke_ != nil {
		this.Handlers.Invoke_(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr)
	}
}

func (this *ApplicationEvents3DispImpl) Startup() {
	if this.Handlers.Startup != nil {
		this.Handlers.Startup()
	}
}

func (this *ApplicationEvents3DispImpl) Quit() {
	if this.Handlers.Quit != nil {
		this.Handlers.Quit()
	}
}

func (this *ApplicationEvents3DispImpl) DocumentChange() {
	if this.Handlers.DocumentChange != nil {
		this.Handlers.DocumentChange()
	}
}

func (this *ApplicationEvents3DispImpl) DocumentOpen(doc *Document) {
	if this.Handlers.DocumentOpen != nil {
		this.Handlers.DocumentOpen(doc)
	}
}

func (this *ApplicationEvents3DispImpl) DocumentBeforeClose(doc *Document, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.DocumentBeforeClose != nil {
		this.Handlers.DocumentBeforeClose(doc, cancel)
	}
}

func (this *ApplicationEvents3DispImpl) DocumentBeforePrint(doc *Document, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.DocumentBeforePrint != nil {
		this.Handlers.DocumentBeforePrint(doc, cancel)
	}
}

func (this *ApplicationEvents3DispImpl) DocumentBeforeSave(doc *Document, saveAsUI *win32.VARIANT_BOOL, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.DocumentBeforeSave != nil {
		this.Handlers.DocumentBeforeSave(doc, saveAsUI, cancel)
	}
}

func (this *ApplicationEvents3DispImpl) NewDocument(doc *Document) {
	if this.Handlers.NewDocument != nil {
		this.Handlers.NewDocument(doc)
	}
}

func (this *ApplicationEvents3DispImpl) WindowActivate(doc *Document, wn *Window) {
	if this.Handlers.WindowActivate != nil {
		this.Handlers.WindowActivate(doc, wn)
	}
}

func (this *ApplicationEvents3DispImpl) WindowDeactivate(doc *Document, wn *Window) {
	if this.Handlers.WindowDeactivate != nil {
		this.Handlers.WindowDeactivate(doc, wn)
	}
}

func (this *ApplicationEvents3DispImpl) WindowSelectionChange(sel *Selection) {
	if this.Handlers.WindowSelectionChange != nil {
		this.Handlers.WindowSelectionChange(sel)
	}
}

func (this *ApplicationEvents3DispImpl) WindowBeforeRightClick(sel *Selection, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.WindowBeforeRightClick != nil {
		this.Handlers.WindowBeforeRightClick(sel, cancel)
	}
}

func (this *ApplicationEvents3DispImpl) WindowBeforeDoubleClick(sel *Selection, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.WindowBeforeDoubleClick != nil {
		this.Handlers.WindowBeforeDoubleClick(sel, cancel)
	}
}

func (this *ApplicationEvents3DispImpl) EPostagePropertyDialog(doc *Document) {
	if this.Handlers.EPostagePropertyDialog != nil {
		this.Handlers.EPostagePropertyDialog(doc)
	}
}

func (this *ApplicationEvents3DispImpl) EPostageInsert(doc *Document) {
	if this.Handlers.EPostageInsert != nil {
		this.Handlers.EPostageInsert(doc)
	}
}

func (this *ApplicationEvents3DispImpl) MailMergeAfterMerge(doc *Document, docResult *Document) {
	if this.Handlers.MailMergeAfterMerge != nil {
		this.Handlers.MailMergeAfterMerge(doc, docResult)
	}
}

func (this *ApplicationEvents3DispImpl) MailMergeAfterRecordMerge(doc *Document) {
	if this.Handlers.MailMergeAfterRecordMerge != nil {
		this.Handlers.MailMergeAfterRecordMerge(doc)
	}
}

func (this *ApplicationEvents3DispImpl) MailMergeBeforeMerge(doc *Document, startRecord int32, endRecord int32, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.MailMergeBeforeMerge != nil {
		this.Handlers.MailMergeBeforeMerge(doc, startRecord, endRecord, cancel)
	}
}

func (this *ApplicationEvents3DispImpl) MailMergeBeforeRecordMerge(doc *Document, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.MailMergeBeforeRecordMerge != nil {
		this.Handlers.MailMergeBeforeRecordMerge(doc, cancel)
	}
}

func (this *ApplicationEvents3DispImpl) MailMergeDataSourceLoad(doc *Document) {
	if this.Handlers.MailMergeDataSourceLoad != nil {
		this.Handlers.MailMergeDataSourceLoad(doc)
	}
}

func (this *ApplicationEvents3DispImpl) MailMergeDataSourceValidate(doc *Document, handled *win32.VARIANT_BOOL) {
	if this.Handlers.MailMergeDataSourceValidate != nil {
		this.Handlers.MailMergeDataSourceValidate(doc, handled)
	}
}

func (this *ApplicationEvents3DispImpl) MailMergeWizardSendToCustom(doc *Document) {
	if this.Handlers.MailMergeWizardSendToCustom != nil {
		this.Handlers.MailMergeWizardSendToCustom(doc)
	}
}

func (this *ApplicationEvents3DispImpl) MailMergeWizardStateChange(doc *Document, fromState *int32, toState *int32, handled *win32.VARIANT_BOOL) {
	if this.Handlers.MailMergeWizardStateChange != nil {
		this.Handlers.MailMergeWizardStateChange(doc, fromState, toState, handled)
	}
}

func (this *ApplicationEvents3DispImpl) WindowSize(doc *Document, wn *Window) {
	if this.Handlers.WindowSize != nil {
		this.Handlers.WindowSize(doc, wn)
	}
}

type ApplicationEvents3Impl struct {
	ole.IDispatchImpl
	DispImpl ApplicationEvents3DispInterface
}

func (this *ApplicationEvents3Impl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == IID_ApplicationEvents3 {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IDispatchImpl.QueryInterface(riid, ppvObject)
}

func (this *ApplicationEvents3Impl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32,
	wFlags uint16, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT,
	pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {
	var unwrapActions ole.Actions
	defer unwrapActions.Execute()
	switch dispIdMember {
	case 1610612736:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*syscall.GUID)(vArgs[0].ToPointer())
		p2 := (unsafe.Pointer)(vArgs[1].ToPointer())
		this.DispImpl.QueryInterface_(p1, p2)
		return win32.S_OK
	case 1610612737:
		ret := this.DispImpl.AddRef_()
		ole.SetVariantParam((*ole.Variant)(pVarResult), ret, &unwrapActions)
		return win32.S_OK
	case 1610612738:
		ret := this.DispImpl.Release_()
		ole.SetVariantParam((*ole.Variant)(pVarResult), ret, &unwrapActions)
		return win32.S_OK
	case 1610678272:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*uint32)(vArgs[0].ToPointer())
		this.DispImpl.GetTypeInfoCount_(p1)
		return win32.S_OK
	case 1610678273:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1, _ := vArgs[0].ToUint32()
		p2, _ := vArgs[1].ToUint32()
		p3 := (unsafe.Pointer)(vArgs[2].ToPointer())
		this.DispImpl.GetTypeInfo_(p1, p2, p3)
		return win32.S_OK
	case 1610678274:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 5)
		p1 := (*syscall.GUID)(vArgs[0].ToPointer())
		p2 := (**int8)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToUint32()
		p4, _ := vArgs[3].ToUint32()
		p5 := (*int32)(vArgs[4].ToPointer())
		this.DispImpl.GetIDsOfNames_(p1, p2, p3, p4, p5)
		return win32.S_OK
	case 1610678275:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 8)
		p1, _ := vArgs[0].ToInt32()
		p2 := (*syscall.GUID)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToUint32()
		p4, _ := vArgs[3].ToUint16()
		p5 := (*win32.DISPPARAMS)(vArgs[4].ToPointer())
		p6 := (*ole.Variant)(vArgs[5].ToPointer())
		p7 := (*win32.EXCEPINFO)(vArgs[6].ToPointer())
		p8 := (*uint32)(vArgs[7].ToPointer())
		this.DispImpl.Invoke_(p1, p2, p3, p4, p5, p6, p7, p8)
		return win32.S_OK
	case 1:
		this.DispImpl.Startup()
		return win32.S_OK
	case 2:
		this.DispImpl.Quit()
		return win32.S_OK
	case 3:
		this.DispImpl.DocumentChange()
		return win32.S_OK
	case 4:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Document)(vArgs[0].ToPointer())
		this.DispImpl.DocumentOpen(p1)
		return win32.S_OK
	case 6:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.DocumentBeforeClose(p1, p2)
		return win32.S_OK
	case 7:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.DocumentBeforePrint(p1, p2)
		return win32.S_OK
	case 8:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		p3 := (*win32.VARIANT_BOOL)(vArgs[2].ToPointer())
		this.DispImpl.DocumentBeforeSave(p1, p2, p3)
		return win32.S_OK
	case 9:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Document)(vArgs[0].ToPointer())
		this.DispImpl.NewDocument(p1)
		return win32.S_OK
	case 10:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*Window)(vArgs[1].ToPointer())
		this.DispImpl.WindowActivate(p1, p2)
		return win32.S_OK
	case 11:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*Window)(vArgs[1].ToPointer())
		this.DispImpl.WindowDeactivate(p1, p2)
		return win32.S_OK
	case 12:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Selection)(vArgs[0].ToPointer())
		this.DispImpl.WindowSelectionChange(p1)
		return win32.S_OK
	case 13:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Selection)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.WindowBeforeRightClick(p1, p2)
		return win32.S_OK
	case 14:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Selection)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.WindowBeforeDoubleClick(p1, p2)
		return win32.S_OK
	case 15:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Document)(vArgs[0].ToPointer())
		this.DispImpl.EPostagePropertyDialog(p1)
		return win32.S_OK
	case 16:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Document)(vArgs[0].ToPointer())
		this.DispImpl.EPostageInsert(p1)
		return win32.S_OK
	case 17:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*Document)(vArgs[1].ToPointer())
		this.DispImpl.MailMergeAfterMerge(p1, p2)
		return win32.S_OK
	case 18:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Document)(vArgs[0].ToPointer())
		this.DispImpl.MailMergeAfterRecordMerge(p1)
		return win32.S_OK
	case 19:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToInt32()
		p3, _ := vArgs[2].ToInt32()
		p4 := (*win32.VARIANT_BOOL)(vArgs[3].ToPointer())
		this.DispImpl.MailMergeBeforeMerge(p1, p2, p3, p4)
		return win32.S_OK
	case 20:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.MailMergeBeforeRecordMerge(p1, p2)
		return win32.S_OK
	case 21:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Document)(vArgs[0].ToPointer())
		this.DispImpl.MailMergeDataSourceLoad(p1)
		return win32.S_OK
	case 22:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.MailMergeDataSourceValidate(p1, p2)
		return win32.S_OK
	case 23:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*Document)(vArgs[0].ToPointer())
		this.DispImpl.MailMergeWizardSendToCustom(p1)
		return win32.S_OK
	case 24:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 4)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*int32)(vArgs[1].ToPointer())
		p3 := (*int32)(vArgs[2].ToPointer())
		p4 := (*win32.VARIANT_BOOL)(vArgs[3].ToPointer())
		this.DispImpl.MailMergeWizardStateChange(p1, p2, p3, p4)
		return win32.S_OK
	case 25:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*Document)(vArgs[0].ToPointer())
		p2 := (*Window)(vArgs[1].ToPointer())
		this.DispImpl.WindowSize(p1, p2)
		return win32.S_OK
	}
	return win32.E_NOTIMPL
}

type ApplicationEvents3ComObj struct {
	ole.IDispatchComObj
}

func NewApplicationEvents3ComObj(dispImpl ApplicationEvents3DispInterface, scoped bool) *ApplicationEvents3ComObj {
	comObj := com.NewComObj[ApplicationEvents3ComObj](
		&ApplicationEvents3Impl {DispImpl: dispImpl})
	if scoped {
		com.AddToScope(comObj)
	}
	return comObj
}

