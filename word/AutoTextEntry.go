package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020936-0000-0000-C000-000000000046
var IID_AutoTextEntry = syscall.GUID{0x00020936, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AutoTextEntry struct {
	ole.OleClient
}

func NewAutoTextEntry(pDisp *win32.IDispatch, addRef bool, scoped bool) *AutoTextEntry {
	 if pDisp == nil {
		return nil;
	}
	p := &AutoTextEntry{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AutoTextEntryFromVar(v ole.Variant) *AutoTextEntry {
	return NewAutoTextEntry(v.IDispatch(), false, false)
}

func (this *AutoTextEntry) IID() *syscall.GUID {
	return &IID_AutoTextEntry
}

func (this *AutoTextEntry) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AutoTextEntry) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AutoTextEntry) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AutoTextEntry) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AutoTextEntry) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *AutoTextEntry) Name() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AutoTextEntry) SetName(rhs string)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *AutoTextEntry) StyleName() string {
	retVal, _ := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AutoTextEntry) Value() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AutoTextEntry) SetValue(rhs string)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *AutoTextEntry) Delete()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

var AutoTextEntry_Insert_OptArgs= []string{
	"RichText", 
}

func (this *AutoTextEntry) Insert(where *Range, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(AutoTextEntry_Insert_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, []interface{}{where}, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

