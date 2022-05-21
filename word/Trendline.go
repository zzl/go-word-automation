package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 91C46192-3124-4346-A815-10B8873F5A06
var IID_Trendline = syscall.GUID{0x91C46192, 0x3124, 0x4346, 
	[8]byte{0xA8, 0x15, 0x10, 0xB8, 0x87, 0x3F, 0x5A, 0x06}}

type Trendline struct {
	ole.OleClient
}

func NewTrendline(pDisp *win32.IDispatch, addRef bool, scoped bool) *Trendline {
	p := &Trendline{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TrendlineFromVar(v ole.Variant) *Trendline {
	return NewTrendline(v.PdispValVal(), false, false)
}

func (this *Trendline) IID() *syscall.GUID {
	return &IID_Trendline
}

func (this *Trendline) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Trendline) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Trendline) Backward() float64 {
	retVal := this.PropGet(0x000000b9, nil)
	return retVal.DblValVal()
}

func (this *Trendline) SetBackward(rhs float64)  {
	retVal := this.PropPut(0x000000b9, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) Border() *ChartBorder {
	retVal := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *Trendline) ClearFormats() ole.Variant {
	retVal := this.Call(0x00000070, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Trendline) DataLabel() *DataLabel {
	retVal := this.PropGet(0x0000009e, nil)
	return NewDataLabel(retVal.PdispValVal(), false, true)
}

func (this *Trendline) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Trendline) DisplayEquation() bool {
	retVal := this.PropGet(0x000000be, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Trendline) SetDisplayEquation(rhs bool)  {
	retVal := this.PropPut(0x000000be, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) DisplayRSquared() bool {
	retVal := this.PropGet(0x000000bd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Trendline) SetDisplayRSquared(rhs bool)  {
	retVal := this.PropPut(0x000000bd, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) Forward() float64 {
	retVal := this.PropGet(0x000000bf, nil)
	return retVal.DblValVal()
}

func (this *Trendline) SetForward(rhs float64)  {
	retVal := this.PropPut(0x000000bf, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) Index() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Trendline) Intercept() float64 {
	retVal := this.PropGet(0x000000ba, nil)
	return retVal.DblValVal()
}

func (this *Trendline) SetIntercept(rhs float64)  {
	retVal := this.PropPut(0x000000ba, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) InterceptIsAuto() bool {
	retVal := this.PropGet(0x000000bb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Trendline) SetInterceptIsAuto(rhs bool)  {
	retVal := this.PropPut(0x000000bb, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Trendline) SetName(rhs string)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) NameIsAuto() bool {
	retVal := this.PropGet(0x000000bc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Trendline) SetNameIsAuto(rhs bool)  {
	retVal := this.PropPut(0x000000bc, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) Order() int32 {
	retVal := this.PropGet(0x000000c0, nil)
	return retVal.LValVal()
}

func (this *Trendline) SetOrder(rhs int32)  {
	retVal := this.PropPut(0x000000c0, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) Period() int32 {
	retVal := this.PropGet(0x000000b8, nil)
	return retVal.LValVal()
}

func (this *Trendline) SetPeriod(rhs int32)  {
	retVal := this.PropPut(0x000000b8, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Trendline) Type() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Trendline) SetType(rhs int32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) Format() *ChartFormat {
	retVal := this.PropGet(0x6002001d, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *Trendline) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Trendline) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Trendline) Backward2() float64 {
	retVal := this.PropGet(0x00000a5a, nil)
	return retVal.DblValVal()
}

func (this *Trendline) SetBackward2(rhs float64)  {
	retVal := this.PropPut(0x00000a5a, []interface{}{rhs})
	_= retVal
}

func (this *Trendline) Forward2() float64 {
	retVal := this.PropGet(0x00000a5b, nil)
	return retVal.DblValVal()
}

func (this *Trendline) SetForward2(rhs float64)  {
	retVal := this.PropPut(0x00000a5b, []interface{}{rhs})
	_= retVal
}

