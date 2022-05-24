package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 56AFD330-440C-4F4C-A39C-ED306D084D5F
var IID_PlotArea = syscall.GUID{0x56AFD330, 0x440C, 0x4F4C, 
	[8]byte{0xA3, 0x9C, 0xED, 0x30, 0x6D, 0x08, 0x4D, 0x5F}}

type PlotArea struct {
	ole.OleClient
}

func NewPlotArea(pDisp *win32.IDispatch, addRef bool, scoped bool) *PlotArea {
	 if pDisp == nil {
		return nil;
	}
	p := &PlotArea{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PlotAreaFromVar(v ole.Variant) *PlotArea {
	return NewPlotArea(v.IDispatch(), false, false)
}

func (this *PlotArea) IID() *syscall.GUID {
	return &IID_PlotArea
}

func (this *PlotArea) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PlotArea) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *PlotArea) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PlotArea) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PlotArea) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *PlotArea) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *PlotArea) Height() float64 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.DblValVal()
}

func (this *PlotArea) SetHeight(rhs float64)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *PlotArea) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *PlotArea) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *PlotArea) Left() float64 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.DblValVal()
}

func (this *PlotArea) SetLeft(rhs float64)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *PlotArea) Top() float64 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.DblValVal()
}

func (this *PlotArea) SetTop(rhs float64)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *PlotArea) Width() float64 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.DblValVal()
}

func (this *PlotArea) SetWidth(rhs float64)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *PlotArea) InsideLeft() float64 {
	retVal, _ := this.PropGet(0x00000683, nil)
	return retVal.DblValVal()
}

func (this *PlotArea) SetInsideLeft(rhs float64)  {
	_ = this.PropPut(0x00000683, []interface{}{rhs})
}

func (this *PlotArea) InsideTop() float64 {
	retVal, _ := this.PropGet(0x00000684, nil)
	return retVal.DblValVal()
}

func (this *PlotArea) SetInsideTop(rhs float64)  {
	_ = this.PropPut(0x00000684, []interface{}{rhs})
}

func (this *PlotArea) InsideWidth() float64 {
	retVal, _ := this.PropGet(0x00000685, nil)
	return retVal.DblValVal()
}

func (this *PlotArea) SetInsideWidth(rhs float64)  {
	_ = this.PropPut(0x00000685, []interface{}{rhs})
}

func (this *PlotArea) InsideHeight() float64 {
	retVal, _ := this.PropGet(0x00000686, nil)
	return retVal.DblValVal()
}

func (this *PlotArea) SetInsideHeight(rhs float64)  {
	_ = this.PropPut(0x00000686, []interface{}{rhs})
}

func (this *PlotArea) Position() int32 {
	retVal, _ := this.PropGet(0x00000687, nil)
	return retVal.LValVal()
}

func (this *PlotArea) SetPosition(rhs int32)  {
	_ = this.PropPut(0x00000687, []interface{}{rhs})
}

func (this *PlotArea) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x60020019, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *PlotArea) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PlotArea) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

