package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 18CD5EC8-8B7B-42C8-992A-2A407468642C
var IID_OMathAutoCorrectEntries = syscall.GUID{0x18CD5EC8, 0x8B7B, 0x42C8, 
	[8]byte{0x99, 0x2A, 0x2A, 0x40, 0x74, 0x68, 0x64, 0x2C}}

type OMathAutoCorrectEntries struct {
	ole.OleClient
}

func NewOMathAutoCorrectEntries(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathAutoCorrectEntries {
	p := &OMathAutoCorrectEntries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathAutoCorrectEntriesFromVar(v ole.Variant) *OMathAutoCorrectEntries {
	return NewOMathAutoCorrectEntries(v.PdispValVal(), false, false)
}

func (this *OMathAutoCorrectEntries) IID() *syscall.GUID {
	return &IID_OMathAutoCorrectEntries
}

func (this *OMathAutoCorrectEntries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathAutoCorrectEntries) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathAutoCorrectEntries) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathAutoCorrectEntries) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathAutoCorrectEntries) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OMathAutoCorrectEntries) ForEach(action func(item *OMathAutoCorrectEntry) bool) {
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
		pItem := (*OMathAutoCorrectEntry)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *OMathAutoCorrectEntries) Count() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *OMathAutoCorrectEntries) Item(index *ole.Variant) *OMathAutoCorrectEntry {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewOMathAutoCorrectEntry(retVal.PdispValVal(), false, true)
}

func (this *OMathAutoCorrectEntries) Add(name string, value string) *OMathAutoCorrectEntry {
	retVal := this.Call(0x000000c8, []interface{}{name, value})
	return NewOMathAutoCorrectEntry(retVal.PdispValVal(), false, true)
}

