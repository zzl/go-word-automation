package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020937-0000-0000-C000-000000000046
var IID_AutoTextEntries = syscall.GUID{0x00020937, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AutoTextEntries struct {
	ole.OleClient
}

func NewAutoTextEntries(pDisp *win32.IDispatch, addRef bool, scoped bool) *AutoTextEntries {
	 if pDisp == nil {
		return nil;
	}
	p := &AutoTextEntries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AutoTextEntriesFromVar(v ole.Variant) *AutoTextEntries {
	return NewAutoTextEntries(v.IDispatch(), false, false)
}

func (this *AutoTextEntries) IID() *syscall.GUID {
	return &IID_AutoTextEntries
}

func (this *AutoTextEntries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AutoTextEntries) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AutoTextEntries) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AutoTextEntries) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AutoTextEntries) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *AutoTextEntries) ForEach(action func(item *AutoTextEntry) bool) {
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
		pItem := (*AutoTextEntry)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *AutoTextEntries) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *AutoTextEntries) Item(index *ole.Variant) *AutoTextEntry {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewAutoTextEntry(retVal.IDispatch(), false, true)
}

func (this *AutoTextEntries) Add(name string, range_ *Range) *AutoTextEntry {
	retVal, _ := this.Call(0x00000065, []interface{}{name, range_})
	return NewAutoTextEntry(retVal.IDispatch(), false, true)
}

func (this *AutoTextEntries) AppendToSpike(range_ *Range) *AutoTextEntry {
	retVal, _ := this.Call(0x00000066, []interface{}{range_})
	return NewAutoTextEntry(retVal.IDispatch(), false, true)
}

