package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020948-0000-0000-C000-000000000046
var IID_AutoCorrectEntries = syscall.GUID{0x00020948, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AutoCorrectEntries struct {
	ole.OleClient
}

func NewAutoCorrectEntries(pDisp *win32.IDispatch, addRef bool, scoped bool) *AutoCorrectEntries {
	p := &AutoCorrectEntries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AutoCorrectEntriesFromVar(v ole.Variant) *AutoCorrectEntries {
	return NewAutoCorrectEntries(v.PdispValVal(), false, false)
}

func (this *AutoCorrectEntries) IID() *syscall.GUID {
	return &IID_AutoCorrectEntries
}

func (this *AutoCorrectEntries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AutoCorrectEntries) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrectEntries) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AutoCorrectEntries) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *AutoCorrectEntries) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *AutoCorrectEntries) ForEach(action func(item *AutoCorrectEntry) bool) {
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
		pItem := (*AutoCorrectEntry)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *AutoCorrectEntries) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *AutoCorrectEntries) Item(index *ole.Variant) *AutoCorrectEntry {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewAutoCorrectEntry(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrectEntries) Add(name string, value string) *AutoCorrectEntry {
	retVal := this.Call(0x00000065, []interface{}{name, value})
	return NewAutoCorrectEntry(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrectEntries) AddRichText(name string, range_ *Range) *AutoCorrectEntry {
	retVal := this.Call(0x00000066, []interface{}{name, range_})
	return NewAutoCorrectEntry(retVal.PdispValVal(), false, true)
}

