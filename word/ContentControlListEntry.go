package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0C6FA8CA-E65F-4FC7-AB8F-20729EECBB14
var IID_ContentControlListEntry = syscall.GUID{0x0C6FA8CA, 0xE65F, 0x4FC7, 
	[8]byte{0xAB, 0x8F, 0x20, 0x72, 0x9E, 0xEC, 0xBB, 0x14}}

type ContentControlListEntry struct {
	ole.OleClient
}

func NewContentControlListEntry(pDisp *win32.IDispatch, addRef bool, scoped bool) *ContentControlListEntry {
	 if pDisp == nil {
		return nil;
	}
	p := &ContentControlListEntry{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ContentControlListEntryFromVar(v ole.Variant) *ContentControlListEntry {
	return NewContentControlListEntry(v.IDispatch(), false, false)
}

func (this *ContentControlListEntry) IID() *syscall.GUID {
	return &IID_ContentControlListEntry
}

func (this *ContentControlListEntry) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ContentControlListEntry) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ContentControlListEntry) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *ContentControlListEntry) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ContentControlListEntry) Text() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControlListEntry) SetText(rhs string)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *ContentControlListEntry) Value() string {
	retVal, _ := this.PropGet(0x00000068, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControlListEntry) SetValue(rhs string)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *ContentControlListEntry) Index() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *ContentControlListEntry) SetIndex(rhs int32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *ContentControlListEntry) Delete()  {
	retVal, _ := this.Call(0x0000006a, nil)
	_= retVal
}

func (this *ContentControlListEntry) MoveUp()  {
	retVal, _ := this.Call(0x0000006b, nil)
	_= retVal
}

func (this *ContentControlListEntry) MoveDown()  {
	retVal, _ := this.Call(0x0000006c, nil)
	_= retVal
}

func (this *ContentControlListEntry) Select()  {
	retVal, _ := this.Call(0x0000006d, nil)
	_= retVal
}

