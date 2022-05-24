package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 1FD94DF1-3569-4465-94FF-E8B22D28EEB0
var IID_DataLabel = syscall.GUID{0x1FD94DF1, 0x3569, 0x4465, 
	[8]byte{0x94, 0xFF, 0xE8, 0xB2, 0x2D, 0x28, 0xEE, 0xB0}}

type DataLabel struct {
	ole.OleClient
}

func NewDataLabel(pDisp *win32.IDispatch, addRef bool, scoped bool) *DataLabel {
	 if pDisp == nil {
		return nil;
	}
	p := &DataLabel{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DataLabelFromVar(v ole.Variant) *DataLabel {
	return NewDataLabel(v.IDispatch(), false, false)
}

func (this *DataLabel) IID() *syscall.GUID {
	return &IID_DataLabel
}

func (this *DataLabel) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DataLabel) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DataLabel) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabel) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabel) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *DataLabel) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabel) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *DataLabel) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *DataLabel) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabel) SetCaption(rhs string)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

var DataLabel_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *DataLabel) Characters(optArgs ...interface{}) *ChartCharacters {
	optArgs = ole.ProcessOptArgs(DataLabel_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000025b, nil, optArgs...)
	return NewChartCharacters(retVal.IDispatch(), false, true)
}

func (this *DataLabel) Font() *ChartFont {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

func (this *DataLabel) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000088, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabel) SetHorizontalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *DataLabel) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *DataLabel) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *DataLabel) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x00000086, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabel) SetOrientation(rhs interface{})  {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *DataLabel) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabel) SetShadow(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *DataLabel) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabel) SetText(rhs string)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *DataLabel) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *DataLabel) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *DataLabel) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x00000089, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabel) SetVerticalAlignment(rhs interface{})  {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *DataLabel) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *DataLabel) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x000003cf, []interface{}{rhs})
}

func (this *DataLabel) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x000005f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabel) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x000005f5, []interface{}{rhs})
}

func (this *DataLabel) AutoText() bool {
	retVal, _ := this.PropGet(0x00000087, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabel) SetAutoText(rhs bool)  {
	_ = this.PropPut(0x00000087, []interface{}{rhs})
}

func (this *DataLabel) NumberFormat() string {
	retVal, _ := this.PropGet(0x000000c1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabel) SetNumberFormat(rhs string)  {
	_ = this.PropPut(0x000000c1, []interface{}{rhs})
}

func (this *DataLabel) NumberFormatLinked() bool {
	retVal, _ := this.PropGet(0x000000c2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabel) SetNumberFormatLinked(rhs bool)  {
	_ = this.PropPut(0x000000c2, []interface{}{rhs})
}

func (this *DataLabel) NumberFormatLocal() ole.Variant {
	retVal, _ := this.PropGet(0x00000449, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabel) SetNumberFormatLocal(rhs interface{})  {
	_ = this.PropPut(0x00000449, []interface{}{rhs})
}

func (this *DataLabel) ShowLegendKey() bool {
	retVal, _ := this.PropGet(0x000000ab, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabel) SetShowLegendKey(rhs bool)  {
	_ = this.PropPut(0x000000ab, []interface{}{rhs})
}

func (this *DataLabel) Type() ole.Variant {
	retVal, _ := this.PropGet(0x0000006c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabel) SetType(rhs interface{})  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *DataLabel) Position() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *DataLabel) SetPosition(rhs int32)  {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *DataLabel) ShowSeriesName() bool {
	retVal, _ := this.PropGet(0x000007e6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabel) SetShowSeriesName(rhs bool)  {
	_ = this.PropPut(0x000007e6, []interface{}{rhs})
}

func (this *DataLabel) ShowCategoryName() bool {
	retVal, _ := this.PropGet(0x000007e7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabel) SetShowCategoryName(rhs bool)  {
	_ = this.PropPut(0x000007e7, []interface{}{rhs})
}

func (this *DataLabel) ShowValue() bool {
	retVal, _ := this.PropGet(0x000007e8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabel) SetShowValue(rhs bool)  {
	_ = this.PropPut(0x000007e8, []interface{}{rhs})
}

func (this *DataLabel) ShowPercentage() bool {
	retVal, _ := this.PropGet(0x000007e9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabel) SetShowPercentage(rhs bool)  {
	_ = this.PropPut(0x000007e9, []interface{}{rhs})
}

func (this *DataLabel) ShowBubbleSize() bool {
	retVal, _ := this.PropGet(0x000007ea, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabel) SetShowBubbleSize(rhs bool)  {
	_ = this.PropPut(0x000007ea, []interface{}{rhs})
}

func (this *DataLabel) Separator() ole.Variant {
	retVal, _ := this.PropGet(0x000007eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataLabel) SetSeparator(rhs interface{})  {
	_ = this.PropPut(0x000007eb, []interface{}{rhs})
}

func (this *DataLabel) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x60020037, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *DataLabel) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DataLabel) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DataLabel) Height() float64 {
	retVal, _ := this.PropGet(0x60020038, nil)
	return retVal.DblValVal()
}

func (this *DataLabel) Width() float64 {
	retVal, _ := this.PropGet(0x000007ee, nil)
	return retVal.DblValVal()
}

func (this *DataLabel) Formula() string {
	retVal, _ := this.PropGet(0x000007ef, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabel) SetFormula(rhs string)  {
	_ = this.PropPut(0x000007ef, []interface{}{rhs})
}

func (this *DataLabel) FormulaR1C1() string {
	retVal, _ := this.PropGet(0x000007f0, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabel) SetFormulaR1C1(rhs string)  {
	_ = this.PropPut(0x000007f0, []interface{}{rhs})
}

func (this *DataLabel) FormulaLocal() string {
	retVal, _ := this.PropGet(0x000007f1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabel) SetFormulaLocal(rhs string)  {
	_ = this.PropPut(0x000007f1, []interface{}{rhs})
}

func (this *DataLabel) FormulaR1C1Local() string {
	retVal, _ := this.PropGet(0x000007f2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabel) SetFormulaR1C1Local(rhs string)  {
	_ = this.PropPut(0x000007f2, []interface{}{rhs})
}

