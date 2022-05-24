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
	 if pDisp == nil {
		return nil;
	}
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
	return NewChartArea(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartArea) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartArea) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartArea) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *ChartArea) Clear() ole.Variant {
	retVal, _ := this.Call(0x0000006f, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartArea) ClearContents() ole.Variant {
	retVal, _ := this.Call(0x00000071, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartArea) Copy() ole.Variant {
	retVal, _ := this.Call(0x00000227, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartArea) Font() *ChartFont {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

func (this *ChartArea) Shadow() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartArea) SetShadow(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *ChartArea) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartArea) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *ChartArea) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *ChartArea) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *ChartArea) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *ChartArea) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *ChartArea) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *ChartArea) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *ChartArea) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *ChartArea) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *ChartArea) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *ChartArea) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x000005f5, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartArea) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x000005f5, []interface{}{rhs})
}

func (this *ChartArea) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x60020017, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *ChartArea) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartArea) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

