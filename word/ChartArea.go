package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// C75AD98A-74E9-49FE-8BF1-544839CC08A5
var IID_ChartArea = syscall.GUID{0xC75AD98A, 0x74E9, 0x49FE, 
	[8]byte{0x8B, 0xF1, 0x54, 0x48, 0x39, 0xCC, 0x08, 0xA5}}

type ChartArea struct {
	ole.OleClient
}

func NewChartArea(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartArea {
	p := &ChartArea{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartAreaFromVar(v ole.Variant) *ChartArea {
	return NewChartArea(v.PdispValVal(), false, false)
}

func (this *ChartArea) IID() *syscall.GUID {
	return &IID_ChartArea
}

func (this *ChartArea) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartArea) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartArea) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartArea) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartArea) Border() *ChartBorder {
	retVal := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *ChartArea) Clear() ole.Variant {
	retVal := this.Call(0x0000006f, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartArea) ClearContents() ole.Variant {
	retVal := this.Call(0x00000071, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartArea) Copy() ole.Variant {
	retVal := this.Call(0x00000227, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartArea) Font() *ChartFont {
	retVal := this.PropGet(0x00000092, nil)
	return NewChartFont(retVal.PdispValVal(), false, true)
}

func (this *ChartArea) Shadow() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartArea) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *ChartArea) ClearFormats() ole.Variant {
	retVal := this.Call(0x00000070, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartArea) Height() float64 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *ChartArea) SetHeight(rhs float64)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *ChartArea) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *ChartArea) Fill() *ChartFillFormat {
	retVal := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartArea) Left() float64 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *ChartArea) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *ChartArea) Top() float64 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *ChartArea) SetTop(rhs float64)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *ChartArea) Width() float64 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *ChartArea) SetWidth(rhs float64)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *ChartArea) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x000005f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartArea) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x000005f5, []interface{}{rhs})
	_= retVal
}

func (this *ChartArea) Format() *ChartFormat {
	retVal := this.PropGet(0x60020017, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartArea) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartArea) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

