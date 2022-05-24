package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// C1AD33E4-F088-40A9-9D2F-D94017D115C4
var IID_ChartTitle = syscall.GUID{0xC1AD33E4, 0xF088, 0x40A9, 
	[8]byte{0x9D, 0x2F, 0xD9, 0x40, 0x17, 0xD1, 0x15, 0xC4}}

type ChartTitle struct {
	ole.OleClient
}

func NewChartTitle(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartTitle {
	 if pDisp == nil {
		return nil;
	}
	p := &ChartTitle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartTitleFromVar(v ole.Variant) *ChartTitle {
	return NewChartTitle(v.IDispatch(), false, false)
}

func (this *ChartTitle) IID() *syscall.GUID {
	return &IID_ChartTitle
}

func (this *ChartTitle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartTitle) Caption() string {
	retVal, _ := this.PropGet(0x60020000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetCaption(rhs string)  {
	_ = this.PropPut(0x60020000, []interface{}{rhs})
}

var ChartTitle_Characters_OptArgs= []string{
	"Start", "Length", 
}

func (this *ChartTitle) Characters(optArgs ...interface{}) *ChartCharacters {
	optArgs = ole.ProcessOptArgs(ChartTitle_Characters_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x60020002, nil, optArgs...)
	return NewChartCharacters(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Font() *ChartFont {
	retVal, _ := this.PropGet(0x60020003, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) HorizontalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x60020004, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) SetHorizontalAlignment(rhs interface{})  {
	_ = this.PropPut(0x60020004, []interface{}{rhs})
}

func (this *ChartTitle) Left() float64 {
	retVal, _ := this.PropGet(0x60020006, nil)
	return retVal.DblValVal()
}

func (this *ChartTitle) SetLeft(rhs float64)  {
	_ = this.PropPut(0x60020006, []interface{}{rhs})
}

func (this *ChartTitle) Orientation() ole.Variant {
	retVal, _ := this.PropGet(0x60020008, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) SetOrientation(rhs interface{})  {
	_ = this.PropPut(0x60020008, []interface{}{rhs})
}

func (this *ChartTitle) Shadow() bool {
	retVal, _ := this.PropGet(0x6002000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartTitle) SetShadow(rhs bool)  {
	_ = this.PropPut(0x6002000a, []interface{}{rhs})
}

func (this *ChartTitle) Text() string {
	retVal, _ := this.PropGet(0x6002000c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetText(rhs string)  {
	_ = this.PropPut(0x6002000c, []interface{}{rhs})
}

func (this *ChartTitle) Top() float64 {
	retVal, _ := this.PropGet(0x6002000e, nil)
	return retVal.DblValVal()
}

func (this *ChartTitle) SetTop(rhs float64)  {
	_ = this.PropPut(0x6002000e, []interface{}{rhs})
}

func (this *ChartTitle) VerticalAlignment() ole.Variant {
	retVal, _ := this.PropGet(0x60020010, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) SetVerticalAlignment(rhs interface{})  {
	_ = this.PropPut(0x60020010, []interface{}{rhs})
}

func (this *ChartTitle) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x60020012, nil)
	return retVal.LValVal()
}

func (this *ChartTitle) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x60020012, []interface{}{rhs})
}

func (this *ChartTitle) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x60020014, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x60020014, []interface{}{rhs})
}

func (this *ChartTitle) Interior() *Interior {
	retVal, _ := this.PropGet(0x60020016, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x60020017, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Delete() ole.Variant {
	retVal, _ := this.Call(0x60020018, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x60020019, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Name() string {
	retVal, _ := this.PropGet(0x6002001a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x6002001b, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartTitle) Select() ole.Variant {
	retVal, _ := this.Call(0x6002001c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartTitle) IncludeInLayout() bool {
	retVal, _ := this.PropGet(0x00000972, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartTitle) SetIncludeInLayout(rhs bool)  {
	_ = this.PropPut(0x00000972, []interface{}{rhs})
}

func (this *ChartTitle) Position() int32 {
	retVal, _ := this.PropGet(0x00000687, nil)
	return retVal.LValVal()
}

func (this *ChartTitle) SetPosition(rhs int32)  {
	_ = this.PropPut(0x00000687, []interface{}{rhs})
}

func (this *ChartTitle) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x60020021, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *ChartTitle) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartTitle) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartTitle) Height() float64 {
	retVal, _ := this.PropGet(0x60020022, nil)
	return retVal.DblValVal()
}

func (this *ChartTitle) Width() float64 {
	retVal, _ := this.PropGet(0x60020025, nil)
	return retVal.DblValVal()
}

func (this *ChartTitle) Formula() string {
	retVal, _ := this.PropGet(0x60020026, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetFormula(rhs string)  {
	_ = this.PropPut(0x60020026, []interface{}{rhs})
}

func (this *ChartTitle) FormulaR1C1() string {
	retVal, _ := this.PropGet(0x60020028, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetFormulaR1C1(rhs string)  {
	_ = this.PropPut(0x60020028, []interface{}{rhs})
}

func (this *ChartTitle) FormulaLocal() string {
	retVal, _ := this.PropGet(0x6002002a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetFormulaLocal(rhs string)  {
	_ = this.PropPut(0x6002002a, []interface{}{rhs})
}

func (this *ChartTitle) FormulaR1C1Local() string {
	retVal, _ := this.PropGet(0x6002002c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartTitle) SetFormulaR1C1Local(rhs string)  {
	_ = this.PropPut(0x6002002c, []interface{}{rhs})
}

