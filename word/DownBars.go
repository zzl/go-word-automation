package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 84A6A663-AEF4-4FCD-83FD-9BB707F157CA
var IID_DownBars = syscall.GUID{0x84A6A663, 0xAEF4, 0x4FCD, 
	[8]byte{0x83, 0xFD, 0x9B, 0xB7, 0x07, 0xF1, 0x57, 0xCA}}

type DownBars struct {
	ole.OleClient
}

func NewDownBars(pDisp *win32.IDispatch, addRef bool, scoped bool) *DownBars {
	p := &DownBars{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DownBarsFromVar(v ole.Variant) *DownBars {
	return NewDownBars(v.PdispValVal(), false, false)
}

func (this *DownBars) IID() *syscall.GUID {
	return &IID_DownBars
}

func (this *DownBars) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DownBars) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DownBars) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DownBars) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DownBars) Border() *ChartBorder {
	retVal := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *DownBars) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DownBars) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *DownBars) Fill() *ChartFillFormat {
	retVal := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.PdispValVal(), false, true)
}

func (this *DownBars) Format() *ChartFormat {
	retVal := this.PropGet(0x60020007, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *DownBars) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DownBars) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

