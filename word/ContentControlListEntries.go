package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 54F46DC4-F6A6-48CC-BD66-46C1DDEADD22
var IID_ContentControlListEntries = syscall.GUID{0x54F46DC4, 0xF6A6, 0x48CC, 
	[8]byte{0xBD, 0x66, 0x46, 0xC1, 0xDD, 0xEA, 0xDD, 0x22}}

type ContentControlListEntries struct {
	ole.OleClient
}

func NewContentControlListEntries(pDisp *win32.IDispatch, addRef bool, scoped bool) *ContentControlListEntries {
	 if pDisp == nil {
		return nil;
	}
	p := &ContentControlListEntries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ContentControlListEntriesFromVar(v ole.Variant) *ContentControlListEntries {
	return NewContentControlListEntries(v.IDispatch(), false, false)
}

func (this *ContentControlListEntries) IID() *syscall.GUID {
	return &IID_ContentControlListEntries
}

func (this *ContentControlListEntries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ContentControlListEntries) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ContentControlListEntries) ForEach(action func(item *ContentControlListEntry) bool) {
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
		pItem := (*ContentControlListEntry)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ContentControlListEntries) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ContentControlListEntries) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *ContentControlListEntries) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ContentControlListEntries) Count() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *ContentControlListEntries) Clear()  {
	retVal, _ := this.Call(0x00000068, nil)
	_= retVal
}

func (this *ContentControlListEntries) Item(index int32) *ContentControlListEntry {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewContentControlListEntry(retVal.IDispatch(), false, true)
}

var ContentControlListEntries_Add_OptArgs= []string{
	"Value", "Index", 
}

func (this *ContentControlListEntries) Add(text string, optArgs ...interface{}) *ContentControlListEntry {
	optArgs = ole.ProcessOptArgs(ContentControlListEntries_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006a, []interface{}{text}, optArgs...)
	return NewContentControlListEntry(retVal.IDispatch(), false, true)
}

