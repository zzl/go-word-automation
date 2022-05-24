package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020924-0000-0000-C000-000000000046
var IID_ListEntries = syscall.GUID{0x00020924, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListEntries struct {
	ole.OleClient
}

func NewListEntries(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListEntries {
	 if pDisp == nil {
		return nil;
	}
	p := &ListEntries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListEntriesFromVar(v ole.Variant) *ListEntries {
	return NewListEntries(v.IDispatch(), false, false)
}

func (this *ListEntries) IID() *syscall.GUID {
	return &IID_ListEntries
}

func (this *ListEntries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListEntries) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListEntries) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ListEntries) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListEntries) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ListEntries) ForEach(action func(item *ListEntry) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*ListEntry)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ListEntries) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *ListEntries) Item(index *ole.Variant) *ListEntry {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewListEntry(retVal.IDispatch(), false, true)
}

var ListEntries_Add_OptArgs= []string{
	"Index", 
}

func (this *ListEntries) Add(name string, optArgs ...interface{}) *ListEntry {
	optArgs = ole.ProcessOptArgs(ListEntries_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, []interface{}{name}, optArgs...)
	return NewListEntry(retVal.IDispatch(), false, true)
}

func (this *ListEntries) Clear()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

