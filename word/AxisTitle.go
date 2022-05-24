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
	 if pDisp == nil {
		return nil;
	}
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
	return NewAxisTitle(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x60020000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetCaption(rhs string)  {
	_ = this.PropPut(0x60020000, []interface{}{rhs})
}

var AxisTitle_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *AxisTitle) Characters(optArgs ...interface{}) *ChartCharacters {
	optArgs = ole.ProcessOptArgs(AxisTitle_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x60020002, nil, optArgs...)
	return NewChartCharacters(retVal.IDispatch(), false, true)
}

func (this *AxisTitle) Font() *ChartFont {
	retVal, _ := this.PropGet(0x60020003, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

func (this *AxisTitle) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x60020004, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *AxisTitle) SetHorizontalAlignment(rhs interface{})  {
	_ = this.PropPut(0x60020004, []interface{}{rhs})
}

func (this *AxisTitle) Left() float64 {
	retVal, _ := this.PropGet(0x60020006, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) SetLeft(rhs float64)  {
	_ = this.PropPut(0x60020006, []interface{}{rhs})
}

func (this *AxisTitle) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x60020008, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *AxisTitle) SetOrientation(rhs interface{})  {
	_ = this.PropPut(0x60020008, []interface{}{rhs})
}

func (this *AxisTitle) Shadow() bool {
	retVal, _ := this.PropGet(0x6002000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AxisTitle) SetShadow(rhs bool)  {
	_ = this.PropPut(0x6002000a, []interface{}{rhs})
}

func (this *AxisTitle) Text() string {
	retVal, _ := this.PropGet(0x6002000c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetText(rhs string)  {
	_ = this.PropPut(0x6002000c, []interface{}{rhs})
}

func (this *AxisTitle) Top() float64 {
	retVal, _ := this.PropGet(0x6002000e, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) SetTop(rhs float64)  {
	_ = this.PropPut(0x6002000e, []interface{}{rhs})
}

func (this *AxisTitle) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x60020010, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *AxisTitle) SetVerticalAlignment(rhs interface{})  {
	_ = this.PropPut(0x60020010, []interface{}{rhs})
}

func (this *AxisTitle) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x60020012, nil)
	return retVal.LValVal()
}

func (this *AxisTitle) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x60020012, []interface{}{rhs})
}

func (this *AxisTitle) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x60020014, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *AxisTitle) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x60020014, []interface{}{rhs})
}

func (this *AxisTitle) Interior() *Interior {
	retVal, _ := this.PropGet(0x60020016, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *AxisTitle) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x60020017, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *AxisTitle) Delete() ole.Variant {
	retVal, _ := this.Call(0x60020018, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *AxisTitle) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x60020019, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *AxisTitle) Name() string {
	retVal, _ := this.PropGet(0x6002001a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x6002001b, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AxisTitle) Select() ole.Variant {
	retVal, _ := this.Call(0x6002001c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *AxisTitle) IncludeInLayout() bool {
	retVal, _ := this.PropGet(0x00000972, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AxisTitle) SetIncludeInLayout(rhs bool)  {
	_ = this.PropPut(0x00000972, []interface{}{rhs})
}

func (this *AxisTitle) Position() int32 {
	retVal, _ := this.PropGet(0x00000687, nil)
	return retVal.LValVal()
}

func (this *AxisTitle) SetPosition(rhs int32)  {
	_ = this.PropPut(0x00000687, []interface{}{rhs})
}

func (this *AxisTitle) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x60020021, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *AxisTitle) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AxisTitle) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *AxisTitle) Height() float64 {
	retVal, _ := this.PropGet(0x60020022, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) Width() float64 {
	retVal, _ := this.PropGet(0x60020025, nil)
	return retVal.DblValVal()
}

func (this *AxisTitle) Formula() string {
	retVal, _ := this.PropGet(0x60020026, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormula(rhs string)  {
	_ = this.PropPut(0x60020026, []interface{}{rhs})
}

func (this *AxisTitle) FormulaR1C1() string {
	retVal, _ := this.PropGet(0x60020028, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormulaR1C1(rhs string)  {
	_ = this.PropPut(0x60020028, []interface{}{rhs})
}

func (this *AxisTitle) FormulaLocal() string {
	retVal, _ := this.PropGet(0x6002002a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormulaLocal(rhs string)  {
	_ = this.PropPut(0x6002002a, []interface{}{rhs})
}

func (this *AxisTitle) FormulaR1C1Local() string {
	retVal, _ := this.PropGet(0x6002002c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AxisTitle) SetFormulaR1C1Local(rhs string)  {
	_ = this.PropPut(0x6002002c, []interface{}{rhs})
}

