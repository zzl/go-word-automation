package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// B6511068-70BF-4751-A741-55C1D41AD96F
var IID_LegendEntries = syscall.GUID{0xB6511068, 0x70BF, 0x4751, 
	[8]byte{0xA7, 0x41, 0x55, 0xC1, 0xD4, 0x1A, 0xD9, 0x6F}}

type LegendEntries struct {
	ole.OleClient
}

func NewLegendEntries(pDisp *win32.IDispatch, addRef bool, scoped bool) *LegendEntries {
	 if pDisp == nil {
		return nil;
	}
	p := &LegendEntries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LegendEntriesFromVar(v ole.Variant) *LegendEntries {
	return NewLegendEntries(v.IDispatch(), false, false)
}

func (this *LegendEntries) IID() *syscall.GUID {
	return &IID_LegendEntries
}

func (this *LegendEntries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *LegendEntries) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *LegendEntries) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *LegendEntries) Item(index interface{}) *LegendEntry {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewLegendEntry(retVal.IDispatch(), false, true)
}

func (this *LegendEntries) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *LegendEntries) ForEach(action func(item *LegendEntry) bool) {
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
		pItem := (*LegendEntry)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *LegendEntries) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *LegendEntries) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *LegendEntries) Default_(index interface{}) *LegendEntry {
	retVal, _ := this.Call(0x60020006, []interface{}{index})
	return NewLegendEntry(retVal.IDispatch(), false, true)
}

