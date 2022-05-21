package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 8B0E45DB-3A7B-42EE-9D17-A92AF69B79C1
var IID_AxisTitle = syscall.GUID{0x8B0E45DB, 0x3A7B, 0x42EE, 
	[8]byte{0x9D, 0x17, 0xA9, 0x2A, 0xF6, 0x9B, 0x79, 0xC1}}

type AxisTitle struct {
	ole.OleClient
}

func NewAxisTitle(pDisp *win32.IDispatch, addRef bool, scoped bool) *AxisTitle {
	p := &AxisTitle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AxisTitleFromVar(v ole.Variant) *AxisTitle {
	return NewAxisTitle(v.PdispValVal(), false, false)
}

func (this *AxisTitle) IID() *syscall.GUID {
	return &IID_AxisTitle
}

func (this *AxisTitle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AxisTitle) Caption() string {
	retVal := this.PropGet(0x60020000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetCaption(rhs string)  {
	retVal := this.PropPut(0x60020000, []interface{}{rhs})
	_= retVal
}

var AxisTitle_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *AxisTitle) Characters(optArgs ...interface{}) *ChartCharacters {
	optArgs = ole.ProcessOptArgs(AxisTitle_Characters_OptArgs, optArgs)
	retVal := this.PropGet(0x60020002, nil, optArgs...)
	return NewChartCharacters(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Font() *ChartFont {
	retVal := this.PropGet(0x60020003, nil)
	return NewChartFont(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x60020004, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x60020004, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Left() float64 {
	retVal := this.PropGet(0x60020006, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) SetLeft(rhs float64)  {
	retVal := this.PropPut(0x60020006, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Orientation() ole.Variant {
	retVal := this.PropGet(0x60020008, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x60020008, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Shadow() bool {
	retVal := this.PropGet(0x6002000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AxisTitle) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x6002000a, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Text() string {
	retVal := this.PropGet(0x6002000c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetText(rhs string)  {
	retVal := this.PropPut(0x6002000c, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Top() float64 {
	retVal := this.PropGet(0x6002000e, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) SetTop(rhs float64)  {
	retVal := this.PropPut(0x6002000e, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x60020010, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x60020010, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) ReadingOrder() int32 {
	retVal := this.PropGet(0x60020012, nil)
	return retVal.LValVal()
}

func (this *AxisTitle) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x60020012, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x60020014, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x60020014, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Interior() *Interior {
	retVal := this.PropGet(0x60020016, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Fill() *ChartFillFormat {
	retVal := this.PropGet(0x60020017, nil)
	return NewChartFillFormat(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Delete() ole.Variant {
	retVal := this.Call(0x60020018, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) Border() *ChartBorder {
	retVal := this.PropGet(0x60020019, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Name() string {
	retVal := this.PropGet(0x6002001a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x6002001b, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *AxisTitle) Select() ole.Variant {
	retVal := this.Call(0x6002001c, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *AxisTitle) IncludeInLayout() bool {
	retVal := this.PropGet(0x00000972, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AxisTitle) SetIncludeInLayout(rhs bool)  {
	retVal := this.PropPut(0x00000972, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Position() int32 {
	retVal := this.PropGet(0x00000687, nil)
	return retVal.LValVal()
}

func (this *AxisTitle) SetPosition(rhs int32)  {
	retVal := this.PropPut(0x00000687, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) Format() *ChartFormat {
	retVal := this.PropGet(0x60020021, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *AxisTitle) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *AxisTitle) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *AxisTitle) Height() float64 {
	retVal := this.PropGet(0x60020022, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) Width() float64 {
	retVal := this.PropGet(0x60020025, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) Formula() string {
	retVal := this.PropGet(0x60020026, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormula(rhs string)  {
	retVal := this.PropPut(0x60020026, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) FormulaR1C1() string {
	retVal := this.PropGet(0x60020028, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormulaR1C1(rhs string)  {
	retVal := this.PropPut(0x60020028, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) FormulaLocal() string {
	retVal := this.PropGet(0x6002002a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormulaLocal(rhs string)  {
	retVal := this.PropPut(0x6002002a, []interface{}{rhs})
	_= retVal
}

func (this *AxisTitle) FormulaR1C1Local() string {
	retVal := this.PropGet(0x6002002c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormulaR1C1Local(rhs string)  {
	retVal := this.PropPut(0x6002002c, []interface{}{rhs})
	_= retVal
}

