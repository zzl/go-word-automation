package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"syscall"
	"unsafe"
)

// 00020A01-0001-0000-C000-000000000046
var IID_IApplicationEvents4 = syscall.GUID{0x00020A01, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IApplicationEvents4 struct {
	win32.IDispatch
}

func NewIApplicationEvents4(pUnk *win32.IUnknown, addRef bool, scoped bool) *IApplicationEvents4 {
	p := (*IApplicationEvents4)(unsafe.Pointer(pUnk))
	if addRef {
		pUnk.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func (this *IApplicationEvents4) IID() *syscall.GUID {
	return &IID_IApplicationEvents4
}

func (this *IApplicationEvents4) Startup() com.Error {
	addr := (*this.LpVtbl)[7]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) Quit() com.Error {
	addr := (*this.LpVtbl)[8]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) DocumentChange() com.Error {
	addr := (*this.LpVtbl)[9]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) DocumentOpen(doc *Document) com.Error {
	addr := (*this.LpVtbl)[10]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) DocumentBeforeClose(doc *Document, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[11]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) DocumentBeforePrint(doc *Document, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[12]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) DocumentBeforeSave(doc *Document, saveAsUI *win32.VARIANT_BOOL, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[13]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(saveAsUI)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) NewDocument(doc *Document) com.Error {
	addr := (*this.LpVtbl)[14]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) WindowActivate(doc *Document, wn *Window) com.Error {
	addr := (*this.LpVtbl)[15]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(wn)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) WindowDeactivate(doc *Document, wn *Window) com.Error {
	addr := (*this.LpVtbl)[16]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(wn)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) WindowSelectionChange(sel *Selection) com.Error {
	addr := (*this.LpVtbl)[17]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) WindowBeforeRightClick(sel *Selection, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[18]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sel)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) WindowBeforeDoubleClick(sel *Selection, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[19]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sel)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) EPostagePropertyDialog(doc *Document) com.Error {
	addr := (*this.LpVtbl)[20]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) EPostageInsert(doc *Document) com.Error {
	addr := (*this.LpVtbl)[21]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) MailMergeAfterMerge(doc *Document, docResult *Document) com.Error {
	addr := (*this.LpVtbl)[22]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(docResult)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) MailMergeAfterRecordMerge(doc *Document) com.Error {
	addr := (*this.LpVtbl)[23]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) MailMergeBeforeMerge(doc *Document, startRecord int32, endRecord int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[24]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(startRecord), uintptr(endRecord), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) MailMergeBeforeRecordMerge(doc *Document, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[25]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) MailMergeDataSourceLoad(doc *Document) com.Error {
	addr := (*this.LpVtbl)[26]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) MailMergeDataSourceValidate(doc *Document, handled *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[27]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(handled)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) MailMergeWizardSendToCustom(doc *Document) com.Error {
	addr := (*this.LpVtbl)[28]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) MailMergeWizardStateChange(doc *Document, fromState *int32, toState *int32, handled *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[29]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(fromState)), uintptr(unsafe.Pointer(toState)), uintptr(unsafe.Pointer(handled)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) WindowSize(doc *Document, wn *Window) com.Error {
	addr := (*this.LpVtbl)[30]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(wn)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) XMLSelectionChange(sel *Selection, oldXMLNode *XMLNode, newXMLNode *XMLNode, reason *int32) com.Error {
	addr := (*this.LpVtbl)[31]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(sel)), uintptr(unsafe.Pointer(oldXMLNode)), uintptr(unsafe.Pointer(newXMLNode)), uintptr(unsafe.Pointer(reason)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) XMLValidationError(xmlnode *XMLNode) com.Error {
	addr := (*this.LpVtbl)[32]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(xmlnode)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) DocumentSync(doc *Document, syncEventType int32) com.Error {
	addr := (*this.LpVtbl)[33]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(syncEventType))
	return com.Error(ret)
}

func (this *IApplicationEvents4) EPostageInsertEx(doc *Document, cpDeliveryAddrStart int32, cpDeliveryAddrEnd int32, cpReturnAddrStart int32, cpReturnAddrEnd int32, xaWidth int32, yaHeight int32, bstrPrinterName string, bstrPaperFeed string, fPrint bool, fCancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[34]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(cpDeliveryAddrStart), uintptr(cpDeliveryAddrEnd), uintptr(cpReturnAddrStart), uintptr(cpReturnAddrEnd), uintptr(xaWidth), uintptr(yaHeight), uintptr(win32.StrToPointer(bstrPrinterName)), uintptr(win32.StrToPointer(bstrPaperFeed)), uintptr(*(*uint8)(unsafe.Pointer(&fPrint))), uintptr(unsafe.Pointer(fCancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) MailMergeDataSourceValidate2(doc *Document, handled *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[35]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(doc)), uintptr(unsafe.Pointer(handled)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) ProtectedViewWindowOpen(pvWindow *ProtectedViewWindow) com.Error {
	addr := (*this.LpVtbl)[36]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvWindow)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) ProtectedViewWindowBeforeEdit(pvWindow *ProtectedViewWindow, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[37]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvWindow)), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) ProtectedViewWindowBeforeClose(pvWindow *ProtectedViewWindow, closeReason int32, cancel *win32.VARIANT_BOOL) com.Error {
	addr := (*this.LpVtbl)[38]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvWindow)), uintptr(closeReason), uintptr(unsafe.Pointer(cancel)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) ProtectedViewWindowSize(pvWindow *ProtectedViewWindow) com.Error {
	addr := (*this.LpVtbl)[39]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvWindow)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) ProtectedViewWindowActivate(pvWindow *ProtectedViewWindow) com.Error {
	addr := (*this.LpVtbl)[40]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvWindow)))
	return com.Error(ret)
}

func (this *IApplicationEvents4) ProtectedViewWindowDeactivate(pvWindow *ProtectedViewWindow) com.Error {
	addr := (*this.LpVtbl)[41]
	ret, _, _ := syscall.SyscallN(addr, uintptr(unsafe.Pointer(this)), uintptr(unsafe.Pointer(pvWindow)))
	return com.Error(ret)
}

