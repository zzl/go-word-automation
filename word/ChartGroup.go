package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 86488FB4-9633-4C93-8057-FC1FA7A847AE
var IID_ChartGroup = syscall.GUID{0x86488FB4, 0x9633, 0x4C93, 
	[8]byte{0x80, 0x57, 0xFC, 0x1F, 0xA7, 0xA8, 0x47, 0xAE}}

type ChartGroup struct {
	ole.OleClient
}

func NewChartGroup(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartGroup {
	p := &ChartGroup{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartGroupFromVar(v ole.Variant) *ChartGroup {
	return NewChartGroup(v.PdispValVal(), false, false)
}

func (this *ChartGroup) IID() *syscall.GUID {
	return &IID_ChartGroup
}

func (this *ChartGroup) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartGroup) AxisGroup() int32 {
	retVal := this.PropGet(0x60020000, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetAxisGroup(rhs int32)  {
	retVal := this.PropPut(0x60020000, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) DoughnutHoleSize() int32 {
	retVal := this.PropGet(0x60020002, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetDoughnutHoleSize(rhs int32)  {
	retVal := this.PropPut(0x60020002, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) DownBars() *DownBars {
	retVal := this.PropGet(0x60020004, nil)
	return NewDownBars(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) DropLines() *DropLines {
	retVal := this.PropGet(0x60020005, nil)
	return NewDropLines(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) FirstSliceAngle() int32 {
	retVal := this.PropGet(0x60020006, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetFirstSliceAngle(rhs int32)  {
	retVal := this.PropPut(0x60020006, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) GapWidth() int32 {
	retVal := this.PropGet(0x60020008, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetGapWidth(rhs int32)  {
	retVal := this.PropPut(0x60020008, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasDropLines() bool {
	retVal := this.PropGet(0x6002000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasDropLines(rhs bool)  {
	retVal := this.PropPut(0x6002000a, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasHiLoLines() bool {
	retVal := this.PropGet(0x6002000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasHiLoLines(rhs bool)  {
	retVal := this.PropPut(0x6002000c, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasRadarAxisLabels() bool {
	retVal := this.PropGet(0x6002000e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasRadarAxisLabels(rhs bool)  {
	retVal := this.PropPut(0x6002000e, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasSeriesLines() bool {
	retVal := this.PropGet(0x60020010, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasSeriesLines(rhs bool)  {
	retVal := this.PropPut(0x60020010, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HasUpDownBars() bool {
	retVal := this.PropGet(0x60020012, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHasUpDownBars(rhs bool)  {
	retVal := this.PropPut(0x60020012, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) HiLoLines() *HiLoLines {
	retVal := this.PropGet(0x60020014, nil)
	return NewHiLoLines(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) Index() int32 {
	retVal := this.PropGet(0x60020015, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) Overlap() int32 {
	retVal := this.PropGet(0x60020016, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetOverlap(rhs int32)  {
	retVal := this.PropPut(0x60020016, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) RadarAxisLabels() *TickLabels {
	retVal := this.PropGet(0x60020018, nil)
	return NewTickLabels(retVal.PdispValVal(), false, true)
}

var ChartGroup_SeriesCollection_OptArgs= []string{
	"Index", 
}

func (this *ChartGroup) SeriesCollection(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(ChartGroup_SeriesCollection_OptArgs, optArgs)
	retVal := this.Call(0x60020019, nil, optArgs...)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartGroup) SeriesLines() *SeriesLines {
	retVal := this.PropGet(0x6002001a, nil)
	return NewSeriesLines(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) SubType() int32 {
	retVal := this.PropGet(0x6002001b, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSubType(rhs int32)  {
	retVal := this.PropPut(0x6002001b, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) Type() int32 {
	retVal := this.PropGet(0x6002001d, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetType(rhs int32)  {
	retVal := this.PropPut(0x6002001d, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) UpBars() *UpBars {
	retVal := this.PropGet(0x6002001f, nil)
	return NewUpBars(retVal.PdispValVal(), false, true)
}

func (this *ChartGroup) VaryByCategories() bool {
	retVal := this.PropGet(0x60020020, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetVaryByCategories(rhs bool)  {
	retVal := this.PropPut(0x60020020, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) SizeRepresents() int32 {
	retVal := this.PropGet(0x60020022, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSizeRepresents(rhs int32)  {
	retVal := this.PropPut(0x60020022, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) BubbleScale() int32 {
	retVal := this.PropGet(0x60020024, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetBubbleScale(rhs int32)  {
	retVal := this.PropPut(0x60020024, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) ShowNegativeBubbles() bool {
	retVal := this.PropGet(0x60020026, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetShowNegativeBubbles(rhs bool)  {
	retVal := this.PropPut(0x60020026, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) SplitType() int32 {
	retVal := this.PropGet(0x60020028, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSplitType(rhs int32)  {
	retVal := this.PropPut(0x60020028, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) SplitValue() ole.Variant {
	retVal := this.PropGet(0x6002002a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartGroup) SetSplitValue(rhs interface{})  {
	retVal := this.PropPut(0x6002002a, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) SecondPlotSize() int32 {
	retVal := this.PropGet(0x6002002c, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) SetSecondPlotSize(rhs int32)  {
	retVal := this.PropPut(0x6002002c, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) Has3DShading() bool {
	retVal := this.PropGet(0x6002002e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartGroup) SetHas3DShading(rhs bool)  {
	retVal := this.PropPut(0x6002002e, []interface{}{rhs})
	_= retVal
}

func (this *ChartGroup) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartGroup) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartGroup) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

