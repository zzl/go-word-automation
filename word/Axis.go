package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 7EBC66BD-F788-42C3-91F4-E8C841A69005
var IID_Axis = syscall.GUID{0x7EBC66BD, 0xF788, 0x42C3, 
	[8]byte{0x91, 0xF4, 0xE8, 0xC8, 0x41, 0xA6, 0x90, 0x05}}

type Axis struct {
	ole.OleClient
}

func NewAxis(pDisp *win32.IDispatch, addRef bool, scoped bool) *Axis {
	 if pDisp == nil {
		return nil;
	}
	p := &Axis{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AxisFromVar(v ole.Variant) *Axis {
	return NewAxis(v.IDispatch(), false, false)
}

func (this *Axis) IID() *syscall.GUID {
	return &IID_Axis
}

func (this *Axis) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Axis) AxisBetweenCategories() bool {
	retVal, _ := this.PropGet(0x60020000, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetAxisBetweenCategories(rhs bool)  {
	_ = this.PropPut(0x60020000, []interface{}{rhs})
}

func (this *Axis) AxisGroup() int32 {
	retVal, _ := this.PropGet(0x60020002, nil)
	return retVal.LValVal()
}

func (this *Axis) AxisTitle() *AxisTitle {
	retVal, _ := this.PropGet(0x60020003, nil)
	return NewAxisTitle(retVal.IDispatch(), false, true)
}

func (this *Axis) CategoryNames() ole.Variant {
	retVal, _ := this.PropGet(0x60020004, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Axis) SetCategoryNames(rhs interface{})  {
	_ = this.PropPut(0x60020004, []interface{}{rhs})
}

func (this *Axis) Crosses() int32 {
	retVal, _ := this.PropGet(0x60020006, nil)
	return retVal.LValVal()
}

func (this *Axis) SetCrosses(rhs int32)  {
	_ = this.PropPut(0x60020006, []interface{}{rhs})
}

func (this *Axis) CrossesAt() float64 {
	retVal, _ := this.PropGet(0x60020008, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetCrossesAt(rhs float64)  {
	_ = this.PropPut(0x60020008, []interface{}{rhs})
}

func (this *Axis) Delete() ole.Variant {
	retVal, _ := this.Call(0x6002000a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Axis) HasMajorGridlines() bool {
	retVal, _ := this.PropGet(0x6002000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetHasMajorGridlines(rhs bool)  {
	_ = this.PropPut(0x6002000b, []interface{}{rhs})
}

func (this *Axis) HasMinorGridlines() bool {
	retVal, _ := this.PropGet(0x6002000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetHasMinorGridlines(rhs bool)  {
	_ = this.PropPut(0x6002000d, []interface{}{rhs})
}

func (this *Axis) HasTitle() bool {
	retVal, _ := this.PropGet(0x6002000f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetHasTitle(rhs bool)  {
	_ = this.PropPut(0x6002000f, []interface{}{rhs})
}

func (this *Axis) MajorGridlines() *Gridlines {
	retVal, _ := this.PropGet(0x60020011, nil)
	return NewGridlines(retVal.IDispatch(), false, true)
}

func (this *Axis) MajorTickMark() int32 {
	retVal, _ := this.PropGet(0x60020012, nil)
	return retVal.LValVal()
}

func (this *Axis) SetMajorTickMark(rhs int32)  {
	_ = this.PropPut(0x60020012, []interface{}{rhs})
}

func (this *Axis) MajorUnit() float64 {
	retVal, _ := this.PropGet(0x60020014, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetMajorUnit(rhs float64)  {
	_ = this.PropPut(0x60020014, []interface{}{rhs})
}

func (this *Axis) LogBase() float64 {
	retVal, _ := this.PropGet(0x60020016, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetLogBase(rhs float64)  {
	_ = this.PropPut(0x60020016, []interface{}{rhs})
}

func (this *Axis) TickLabelSpacingIsAuto() bool {
	retVal, _ := this.PropGet(0x60020018, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetTickLabelSpacingIsAuto(rhs bool)  {
	_ = this.PropPut(0x60020018, []interface{}{rhs})
}

func (this *Axis) MajorUnitIsAuto() bool {
	retVal, _ := this.PropGet(0x6002001a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetMajorUnitIsAuto(rhs bool)  {
	_ = this.PropPut(0x6002001a, []interface{}{rhs})
}

func (this *Axis) MaximumScale() float64 {
	retVal, _ := this.PropGet(0x6002001c, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetMaximumScale(rhs float64)  {
	_ = this.PropPut(0x6002001c, []interface{}{rhs})
}

func (this *Axis) MaximumScaleIsAuto() bool {
	retVal, _ := this.PropGet(0x6002001e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetMaximumScaleIsAuto(rhs bool)  {
	_ = this.PropPut(0x6002001e, []interface{}{rhs})
}

func (this *Axis) MinimumScale() float64 {
	retVal, _ := this.PropGet(0x60020020, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetMinimumScale(rhs float64)  {
	_ = this.PropPut(0x60020020, []interface{}{rhs})
}

func (this *Axis) MinimumScaleIsAuto() bool {
	retVal, _ := this.PropGet(0x60020022, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetMinimumScaleIsAuto(rhs bool)  {
	_ = this.PropPut(0x60020022, []interface{}{rhs})
}

func (this *Axis) MinorGridlines() *Gridlines {
	retVal, _ := this.PropGet(0x60020024, nil)
	return NewGridlines(retVal.IDispatch(), false, true)
}

func (this *Axis) MinorTickMark() int32 {
	retVal, _ := this.PropGet(0x60020025, nil)
	return retVal.LValVal()
}

func (this *Axis) SetMinorTickMark(rhs int32)  {
	_ = this.PropPut(0x60020025, []interface{}{rhs})
}

func (this *Axis) MinorUnit() float64 {
	retVal, _ := this.PropGet(0x60020027, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetMinorUnit(rhs float64)  {
	_ = this.PropPut(0x60020027, []interface{}{rhs})
}

func (this *Axis) MinorUnitIsAuto() bool {
	retVal, _ := this.PropGet(0x60020029, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetMinorUnitIsAuto(rhs bool)  {
	_ = this.PropPut(0x60020029, []interface{}{rhs})
}

func (this *Axis) ReversePlotOrder() bool {
	retVal, _ := this.PropGet(0x6002002b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetReversePlotOrder(rhs bool)  {
	_ = this.PropPut(0x6002002b, []interface{}{rhs})
}

func (this *Axis) ScaleType() int32 {
	retVal, _ := this.PropGet(0x6002002d, nil)
	return retVal.LValVal()
}

func (this *Axis) SetScaleType(rhs int32)  {
	_ = this.PropPut(0x6002002d, []interface{}{rhs})
}

func (this *Axis) Select() ole.Variant {
	retVal, _ := this.Call(0x6002002f, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Axis) TickLabelPosition() int32 {
	retVal, _ := this.PropGet(0x60020030, nil)
	return retVal.LValVal()
}

func (this *Axis) SetTickLabelPosition(rhs int32)  {
	_ = this.PropPut(0x60020030, []interface{}{rhs})
}

func (this *Axis) TickLabels() *TickLabels {
	retVal, _ := this.PropGet(0x60020032, nil)
	return NewTickLabels(retVal.IDispatch(), false, true)
}

func (this *Axis) TickLabelSpacing() int32 {
	retVal, _ := this.PropGet(0x60020033, nil)
	return retVal.LValVal()
}

func (this *Axis) SetTickLabelSpacing(rhs int32)  {
	_ = this.PropPut(0x60020033, []interface{}{rhs})
}

func (this *Axis) TickMarkSpacing() int32 {
	retVal, _ := this.PropGet(0x60020035, nil)
	return retVal.LValVal()
}

func (this *Axis) SetTickMarkSpacing(rhs int32)  {
	_ = this.PropPut(0x60020035, []interface{}{rhs})
}

func (this *Axis) Type() int32 {
	retVal, _ := this.PropGet(0x60020037, nil)
	return retVal.LValVal()
}

func (this *Axis) SetType(rhs int32)  {
	_ = this.PropPut(0x60020037, []interface{}{rhs})
}

func (this *Axis) BaseUnit() int32 {
	retVal, _ := this.PropGet(0x60020039, nil)
	return retVal.LValVal()
}

func (this *Axis) SetBaseUnit(rhs int32)  {
	_ = this.PropPut(0x60020039, []interface{}{rhs})
}

func (this *Axis) BaseUnitIsAuto() bool {
	retVal, _ := this.PropGet(0x6002003b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetBaseUnitIsAuto(rhs bool)  {
	_ = this.PropPut(0x6002003b, []interface{}{rhs})
}

func (this *Axis) MajorUnitScale() int32 {
	retVal, _ := this.PropGet(0x6002003d, nil)
	return retVal.LValVal()
}

func (this *Axis) SetMajorUnitScale(rhs int32)  {
	_ = this.PropPut(0x6002003d, []interface{}{rhs})
}

func (this *Axis) MinorUnitScale() int32 {
	retVal, _ := this.PropGet(0x6002003f, nil)
	return retVal.LValVal()
}

func (this *Axis) SetMinorUnitScale(rhs int32)  {
	_ = this.PropPut(0x6002003f, []interface{}{rhs})
}

func (this *Axis) CategoryType() int32 {
	retVal, _ := this.PropGet(0x60020041, nil)
	return retVal.LValVal()
}

func (this *Axis) SetCategoryType(rhs int32)  {
	_ = this.PropPut(0x60020041, []interface{}{rhs})
}

func (this *Axis) Left() float64 {
	retVal, _ := this.PropGet(0x60020043, nil)
	return retVal.DblValVal()
}

func (this *Axis) Top() float64 {
	retVal, _ := this.PropGet(0x60020044, nil)
	return retVal.DblValVal()
}

func (this *Axis) Width() float64 {
	retVal, _ := this.PropGet(0x60020045, nil)
	return retVal.DblValVal()
}

func (this *Axis) Height() float64 {
	retVal, _ := this.PropGet(0x60020046, nil)
	return retVal.DblValVal()
}

func (this *Axis) DisplayUnit() int32 {
	retVal, _ := this.PropGet(0x60020047, nil)
	return retVal.LValVal()
}

func (this *Axis) SetDisplayUnit(rhs int32)  {
	_ = this.PropPut(0x60020047, []interface{}{rhs})
}

func (this *Axis) DisplayUnitCustom() float64 {
	retVal, _ := this.PropGet(0x60020049, nil)
	return retVal.DblValVal()
}

func (this *Axis) SetDisplayUnitCustom(rhs float64)  {
	_ = this.PropPut(0x60020049, []interface{}{rhs})
}

func (this *Axis) HasDisplayUnitLabel() bool {
	retVal, _ := this.PropGet(0x6002004b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Axis) SetHasDisplayUnitLabel(rhs bool)  {
	_ = this.PropPut(0x6002004b, []interface{}{rhs})
}

func (this *Axis) DisplayUnitLabel() *DisplayUnitLabel {
	retVal, _ := this.PropGet(0x6002004d, nil)
	return NewDisplayUnitLabel(retVal.IDispatch(), false, true)
}

func (this *Axis) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x6002004e, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *Axis) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x60020050, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *Axis) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Axis) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Axis) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

