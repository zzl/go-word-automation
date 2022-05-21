package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209FE-0001-0000-C000-000000000046
var IID_IApplicationEvents2 = syscall.GUID{0x000209FE, 0x0001, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type IApplicationEvents2 struct {
	ole.OleClient
}

func NewIApplicationEvents2(pDisp *win32.IDispatch, addRef bool, scoped bool) *IApplicationEvents2 {
	p := &IApplicationEvents2{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func IApplicationEvents2FromVar(v ole.Variant) *IApplicationEvents2 {
	return NewIApplicationEvents2(v.PdispValVal(), false, false)
}

func (this *IApplicationEvents2) IID() *syscall.GUID {
	return &IID_IApplicationEvents2
}

func (this *IApplicationEvents2) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *IApplicationEvents2) Startup()  {
	retVal := this.Call(0x00000001, nil)
	_= retVal
}

func (this *IApplicationEvents2) Quit()  {
	retVal := this.Call(0x00000002, nil)
	_= retVal
}

func (this *IApplicationEvents2) DocumentChange()  {
	retVal := this.Call(0x00000003, nil)
	_= retVal
}

func (this *IApplicationEvents2) DocumentOpen(doc *Document)  {
	retVal := this.Call(0x00000004, []interface{}{doc})
	_= retVal
}

func (this *IApplicationEvents2) DocumentBeforeClose(doc *Document, cancel *win32.VARIANT_BOOL)  {
	retVal := this.Call(0x00000006, []interface{}{doc, cancel})
	_= retVal
}

func (this *IApplicationEvents2) DocumentBeforePrint(doc *Document, cancel *win32.VARIANT_BOOL)  {
	retVal := this.Call(0x00000007, []interface{}{doc, cancel})
	_= retVal
}

func (this *IApplicationEvents2) DocumentBeforeSave(doc *Document, saveAsUI *win32.VARIANT_BOOL, cancel *win32.VARIANT_BOOL)  {
	retVal := this.Call(0x00000008, []interface{}{doc, saveAsUI, cancel})
	_= retVal
}

func (this *IApplicationEvents2) NewDocument(doc *Document)  {
	retVal := this.Call(0x00000009, []interface{}{doc})
	_= retVal
}

func (this *IApplicationEvents2) WindowActivate(doc *Document, wn *Window)  {
	retVal := this.Call(0x0000000a, []interface{}{doc, wn})
	_= retVal
}

func (this *IApplicationEvents2) WindowDeactivate(doc *Document, wn *Window)  {
	retVal := this.Call(0x0000000b, []interface{}{doc, wn})
	_= retVal
}

func (this *IApplicationEvents2) WindowSelectionChange(sel *Selection)  {
	retVal := this.Call(0x0000000c, []interface{}{sel})
	_= retVal
}

func (this *IApplicationEvents2) WindowBeforeRightClick(sel *Selection, cancel *win32.VARIANT_BOOL)  {
	retVal := this.Call(0x0000000d, []interface{}{sel, cancel})
	_= retVal
}

func (this *IApplicationEvents2) WindowBeforeDoubleClick(sel *Selection, cancel *win32.VARIANT_BOOL)  {
	retVal := this.Call(0x0000000e, []interface{}{sel, cancel})
	_= retVal
}

