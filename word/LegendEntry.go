package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// C4A02049-024C-4273-8934-E48CC21479A9
var IID_LegendEntry = syscall.GUID{0xC4A02049, 0x024C, 0x4273, 
	[8]byte{0x89, 0x34, 0xE4, 0x8C, 0xC2, 0x14, 0x79, 0xA9}}

type LegendEntry struct {
	ole.OleClient
}

func NewLegendEntry(pDisp *win32.IDispatch, addRef bool, scoped bool) *LegendEntry {
	 if pDisp == nil {
		return nil;
	}
	p := &LegendEntry{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LegendEntryFromVar(v ole.Variant) *LegendEntry {
	return NewLegendEntry(v.IDispatch(), false, false)
}

func (this *LegendEntry) IID() *syscall.GUID {
	return &IID_LegendEntry
}

func (this *LegendEntry) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *LegendEntry) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *LegendEntry) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *LegendEntry) Font() *ChartFont {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

func (this *LegendEntry) Index() int32 {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *LegendEntry) LegendKey() *LegendKey {
	retVal, _ := this.PropGet(0x000000ae, nil)
	return NewLegendKey(retVal.IDispatch(), false, true)
}

func (this *LegendEntry) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *LegendEntry) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x000005f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *LegendEntry) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x000005f5, []interface{}{rhs})
}

func (this *LegendEntry) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *LegendEntry) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *LegendEntry) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *LegendEntry) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *LegendEntry) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x6002000c, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *LegendEntry) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *LegendEntry) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

