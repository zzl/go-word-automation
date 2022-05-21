package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 86905AC9-33F3-4A88-96C8-B289B0390BCA
var IID_UpBars = syscall.GUID{0x86905AC9, 0x33F3, 0x4A88, 
	[8]byte{0x96, 0xC8, 0xB2, 0x89, 0xB0, 0x39, 0x0B, 0xCA}}

type UpBars struct {
	ole.OleClient
}

func NewUpBars(pDisp *win32.IDispatch, addRef bool, scoped bool) *UpBars {
	p := &UpBars{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func UpBarsFromVar(v ole.Variant) *UpBars {
	return NewUpBars(v.PdispValVal(), false, false)
}

func (this *UpBars) IID() *syscall.GUID {
	return &IID_UpBars
}

func (this *UpBars) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *UpBars) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *UpBars) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *UpBars) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *UpBars) Border() *ChartBorder {
	retVal := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *UpBars) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *UpBars) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *UpBars) Fill() *ChartFillFormat {
	retVal := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.PdispValVal(), false, true)
}

func (this *UpBars) Format() *ChartFormat {
	retVal := this.PropGet(0x60020007, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *UpBars) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *UpBars) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

