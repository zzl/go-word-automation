package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020923-0000-0000-C000-000000000046
var IID_ListEntry = syscall.GUID{0x00020923, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListEntry struct {
	ole.OleClient
}

func NewListEntry(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListEntry {
	 if pDisp == nil {
		return nil;
	}
	p := &ListEntry{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListEntryFromVar(v ole.Variant) *ListEntry {
	return NewListEntry(v.IDispatch(), false, false)
}

func (this *ListEntry) IID() *syscall.GUID {
	return &IID_ListEntry
}

func (this *ListEntry) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListEntry) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListEntry) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ListEntry) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListEntry) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *ListEntry) Name() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListEntry) SetName(rhs string)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *ListEntry) Delete()  {
	retVal, _ := this.Call(0x0000000b, nil)
	_= retVal
}

