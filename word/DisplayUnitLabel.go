package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// C04865A3-9F8A-486C-BB58-B4C3E6563136
var IID_DisplayUnitLabel = syscall.GUID{0xC04865A3, 0x9F8A, 0x486C, 
	[8]byte{0xBB, 0x58, 0xB4, 0xC3, 0xE6, 0x56, 0x31, 0x36}}

type DisplayUnitLabel struct {
	ole.OleClient
}

func NewDisplayUnitLabel(pDisp *win32.IDispatch, addRef bool, scoped bool) *DisplayUnitLabel {
	 if pDisp == nil {
		return nil;
	}
	p := &DisplayUnitLabel{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DisplayUnitLabelFromVar(v ole.Variant) *DisplayUnitLabel {
	return NewDisplayUnitLabel(v.IDispatch(), false, false)
}

func (this *DisplayUnitLabel) IID() *syscall.GUID {
	return &IID_DisplayUnitLabel
}

func (this *DisplayUnitLabel) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DisplayUnitLabel) Caption() string {
	retVal, _ := this.PropGet(0x60020000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetCaption(rhs string)  {
	_ = this.PropPut(0x60020000, []interface{}{rhs})
}

var DisplayUnitLabel_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *DisplayUnitLabel) Characters(optArgs ...interface{}) *ChartCharacters {
	optArgs = ole.ProcessOptArgs(DisplayUnitLabel_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x60020002, nil, optArgs...)
	return NewChartCharacters(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Font() *ChartFont {
	retVal, _ := this.PropGet(0x60020003, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x60020004, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) SetHorizontalAlignment(rhs interface{})  {
	_ = this.PropPut(0x60020004, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Left() float64 {
	retVal, _ := this.PropGet(0x60020006, nil)
	return retVal.DblValVal()
}

func (this *DisplayUnitLabel) SetLeft(rhs float64)  {
	_ = this.PropPut(0x60020006, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x60020008, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) SetOrientation(rhs interface{})  {
	_ = this.PropPut(0x60020008, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Shadow() bool {
	retVal, _ := this.PropGet(0x6002000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DisplayUnitLabel) SetShadow(rhs bool)  {
	_ = this.PropPut(0x6002000a, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Text() string {
	retVal, _ := this.PropGet(0x6002000c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetText(rhs string)  {
	_ = this.PropPut(0x6002000c, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Top() float64 {
	retVal, _ := this.PropGet(0x6002000e, nil)
	return retVal.DblValVal()
}

func (this *DisplayUnitLabel) SetTop(rhs float64)  {
	_ = this.PropPut(0x6002000e, []interface{}{rhs})
}

func (this *DisplayUnitLabel) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x60020010, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) SetVerticalAlignment(rhs interface{})  {
	_ = this.PropPut(0x60020010, []interface{}{rhs})
}

func (this *DisplayUnitLabel) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x60020012, nil)
	return retVal.LValVal()
}

func (this *DisplayUnitLabel) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x60020012, []interface{}{rhs})
}

func (this *DisplayUnitLabel) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x60020014, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x60020014, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Interior() *Interior {
	retVal, _ := this.PropGet(0x60020016, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x60020017, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Delete() ole.Variant {
	retVal, _ := this.Call(0x60020018, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x60020019, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Name() string {
	retVal, _ := this.PropGet(0x6002001a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x6002001b, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DisplayUnitLabel) Select() ole.Variant {
	retVal, _ := this.Call(0x6002001c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DisplayUnitLabel) IncludeInLayout() bool {
	retVal, _ := this.PropGet(0x00000972, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DisplayUnitLabel) SetIncludeInLayout(rhs bool)  {
	_ = this.PropPut(0x00000972, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Position() int32 {
	retVal, _ := this.PropGet(0x00000687, nil)
	return retVal.LValVal()
}

func (this *DisplayUnitLabel) SetPosition(rhs int32)  {
	_ = this.PropPut(0x00000687, []interface{}{rhs})
}

func (this *DisplayUnitLabel) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x60020021, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *DisplayUnitLabel) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DisplayUnitLabel) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DisplayUnitLabel) Height() float64 {
	retVal, _ := this.PropGet(0x60020022, nil)
	return retVal.DblValVal()
}

func (this *DisplayUnitLabel) Width() float64 {
	retVal, _ := this.PropGet(0x60020025, nil)
	return retVal.DblValVal()
}

func (this *DisplayUnitLabel) Formula() string {
	retVal, _ := this.PropGet(0x60020026, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetFormula(rhs string)  {
	_ = this.PropPut(0x60020026, []interface{}{rhs})
}

func (this *DisplayUnitLabel) FormulaR1C1() string {
	retVal, _ := this.PropGet(0x60020028, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetFormulaR1C1(rhs string)  {
	_ = this.PropPut(0x60020028, []interface{}{rhs})
}

func (this *DisplayUnitLabel) FormulaLocal() string {
	retVal, _ := this.PropGet(0x6002002a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetFormulaLocal(rhs string)  {
	_ = this.PropPut(0x6002002a, []interface{}{rhs})
}

func (this *DisplayUnitLabel) FormulaR1C1Local() string {
	retVal, _ := this.PropGet(0x6002002c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DisplayUnitLabel) SetFormulaR1C1Local(rhs string)  {
	_ = this.PropPut(0x6002002c, []interface{}{rhs})
}

