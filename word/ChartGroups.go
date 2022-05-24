package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// F8DDB497-CA6C-4711-9BA4-2718FA3BB6FE
var IID_ChartGroups = syscall.GUID{0xF8DDB497, 0xCA6C, 0x4711, 
	[8]byte{0x9B, 0xA4, 0x27, 0x18, 0xFA, 0x3B, 0xB6, 0xFE}}

type ChartGroups struct {
	ole.OleClient
}

func NewChartGroups(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartGroups {
	 if pDisp == nil {
		return nil;
	}
	p := &ChartGroups{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartGroupsFromVar(v ole.Variant) *ChartGroups {
	return NewChartGroups(v.IDispatch(), false, false)
}

func (this *ChartGroups) IID() *syscall.GUID {
	return &IID_ChartGroups
}

func (this *ChartGroups) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartGroups) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartGroups) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *ChartGroups) Item(index interface{}) *ChartGroup {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewChartGroup(retVal.IDispatch(), false, true)
}

func (this *ChartGroups) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ChartGroups) ForEach(action func(item *ChartGroup) bool) {
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
		pItem := (*ChartGroup)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ChartGroups) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartGroups) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

