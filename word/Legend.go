package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// B3A1E8C6-E1CE-4A46-8D12-E017157B03D7
var IID_Legend = syscall.GUID{0xB3A1E8C6, 0xE1CE, 0x4A46, 
	[8]byte{0x8D, 0x12, 0xE0, 0x17, 0x15, 0x7B, 0x03, 0xD7}}

type Legend struct {
	ole.OleClient
}

func NewLegend(pDisp *win32.IDispatch, addRef bool, scoped bool) *Legend {
	 if pDisp == nil {
		return nil;
	}
	p := &Legend{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LegendFromVar(v ole.Variant) *Legend {
	return NewLegend(v.IDispatch(), false, false)
}

func (this *Legend) IID() *syscall.GUID {
	return &IID_Legend
}

func (this *Legend) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Legend) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Legend) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Legend) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Legend) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *Legend) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Legend) Font() *ChartFont {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

var Legend_LegendEntries_OptArgs= []string{
	"Index", 
}

func (this *Legend) LegendEntries(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Legend_LegendEntries_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000ad, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Legend) Position() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *Legend) SetPosition(rhs int32)  {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *Legend) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Legend) SetShadow(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Legend) Clear() ole.Variant {
	retVal, _ := this.Call(0x0000006f, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Legend) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *Legend) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Legend) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *Legend) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *Legend) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *Legend) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Legend) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *Legend) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Legend) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *Legend) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *Legend) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x000005f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Legend) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x000005f5, []interface{}{rhs})
}

func (this *Legend) IncludeInLayout() bool {
	retVal, _ := this.PropGet(0x00000972, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Legend) SetIncludeInLayout(rhs bool)  {
	_ = this.PropPut(0x00000972, []interface{}{rhs})
}

func (this *Legend) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x6002001a, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *Legend) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Legend) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

