package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020947-0000-0000-C000-000000000046
var IID_AutoCorrectEntry = syscall.GUID{0x00020947, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AutoCorrectEntry struct {
	ole.OleClient
}

func NewAutoCorrectEntry(pDisp *win32.IDispatch, addRef bool, scoped bool) *AutoCorrectEntry {
	p := &AutoCorrectEntry{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AutoCorrectEntryFromVar(v ole.Variant) *AutoCorrectEntry {
	return NewAutoCorrectEntry(v.PdispValVal(), false, false)
}

func (this *AutoCorrectEntry) IID() *syscall.GUID {
	return &IID_AutoCorrectEntry
}

func (this *AutoCorrectEntry) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AutoCorrectEntry) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrectEntry) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AutoCorrectEntry) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *AutoCorrectEntry) Index() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *AutoCorrectEntry) Name() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AutoCorrectEntry) SetName(rhs string)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrectEntry) Value() string {
	retVal := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AutoCorrectEntry) SetValue(rhs string)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrectEntry) RichText() bool {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrectEntry) Delete()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *AutoCorrectEntry) Apply(range_ *Range)  {
	retVal := this.Call(0x00000066, []interface{}{range_})
	_= retVal
}

