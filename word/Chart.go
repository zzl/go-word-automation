package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 6FFA84BB-A350-4442-BB53-A43653459A84
var IID_Chart = syscall.GUID{0x6FFA84BB, 0xA350, 0x4442, 
	[8]byte{0xBB, 0x53, 0xA4, 0x36, 0x53, 0x45, 0x9A, 0x84}}

type Chart struct {
	ole.OleClient
}

func NewChart(pDisp *win32.IDispatch, addRef bool, scoped bool) *Chart {
	 if pDisp == nil {
		return nil;
	}
	p := &Chart{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartFromVar(v ole.Variant) *Chart {
	return NewChart(v.IDispatch(), false, false)
}

func (this *Chart) IID() *syscall.GUID {
	return &IID_Chart
}

func (this *Chart) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Chart) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) HasTitle() bool {
	retVal, _ := this.PropGet(0x60020001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetHasTitle(rhs bool)  {
	_ = this.PropPut(0x60020001, []interface{}{rhs})
}

func (this *Chart) ChartTitle() *ChartTitle {
	retVal, _ := this.PropGet(0x60020003, nil)
	return NewChartTitle(retVal.IDispatch(), false, true)
}

func (this *Chart) DepthPercent() int32 {
	retVal, _ := this.PropGet(0x60020004, nil)
	return retVal.LValVal()
}

func (this *Chart) SetDepthPercent(rhs int32)  {
	_ = this.PropPut(0x60020004, []interface{}{rhs})
}

func (this *Chart) Elevation() int32 {
	retVal, _ := this.PropGet(0x60020006, nil)
	return retVal.LValVal()
}

func (this *Chart) SetElevation(rhs int32)  {
	_ = this.PropPut(0x60020006, []interface{}{rhs})
}

func (this *Chart) GapDepth() int32 {
	retVal, _ := this.PropGet(0x60020008, nil)
	return retVal.LValVal()
}

func (this *Chart) SetGapDepth(rhs int32)  {
	_ = this.PropPut(0x60020008, []interface{}{rhs})
}

func (this *Chart) HeightPercent() int32 {
	retVal, _ := this.PropGet(0x6002000a, nil)
	return retVal.LValVal()
}

func (this *Chart) SetHeightPercent(rhs int32)  {
	_ = this.PropPut(0x6002000a, []interface{}{rhs})
}

func (this *Chart) Perspective() int32 {
	retVal, _ := this.PropGet(0x6002000c, nil)
	return retVal.LValVal()
}

func (this *Chart) SetPerspective(rhs int32)  {
	_ = this.PropPut(0x6002000c, []interface{}{rhs})
}

func (this *Chart) RightAngleAxes() ole.Variant {
	retVal, _ := this.PropGet(0x6002000e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Chart) SetRightAngleAxes(rhs interface{})  {
	_ = this.PropPut(0x6002000e, []interface{}{rhs})
}

func (this *Chart) Rotation() ole.Variant {
	retVal, _ := this.PropGet(0x60020010, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Chart) SetRotation(rhs interface{})  {
	_ = this.PropPut(0x60020010, []interface{}{rhs})
}

func (this *Chart) DisplayBlanksAs() int32 {
	retVal, _ := this.PropGet(0x60020012, nil)
	return retVal.LValVal()
}

func (this *Chart) SetDisplayBlanksAs(rhs int32)  {
	_ = this.PropPut(0x60020012, []interface{}{rhs})
}

var Chart_ChartGroups_OptArgs= []string{
	"Index", 
}

func (this *Chart) ChartGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_ChartGroups_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000008, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart_SeriesCollection_OptArgs= []string{
	"Index", 
}

func (this *Chart) SeriesCollection(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_SeriesCollection_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000044, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) SubType() int32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *Chart) SetSubType(rhs int32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *Chart) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *Chart) SetType(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *Chart) Corners() *Corners {
	retVal, _ := this.PropGet(0x0000004f, nil)
	return NewCorners(retVal.IDispatch(), false, true)
}

var Chart_ApplyDataLabels_OptArgs= []string{
	"Type", "LegendKey", "AutoText", "HasLeaderLines", 
	"ShowSeriesName", "ShowCategoryName", "ShowValue", "ShowPercentage", 
	"ShowBubbleSize", "Separator", 
}

func (this *Chart) ApplyDataLabels(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_ApplyDataLabels_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000782, nil, optArgs...)
	_= retVal
}

func (this *Chart) ChartType() int32 {
	retVal, _ := this.PropGet(0x00000578, nil)
	return retVal.LValVal()
}

func (this *Chart) SetChartType(rhs int32)  {
	_ = this.PropPut(0x00000578, []interface{}{rhs})
}

func (this *Chart) HasDataTable() bool {
	retVal, _ := this.PropGet(0x00000574, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetHasDataTable(rhs bool)  {
	_ = this.PropPut(0x00000574, []interface{}{rhs})
}

var Chart_ApplyCustomType_OptArgs= []string{
	"TypeName", 
}

func (this *Chart) ApplyCustomType(chartType int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_ApplyCustomType_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000579, []interface{}{chartType}, optArgs...)
	_= retVal
}

func (this *Chart) GetChartElement(x int32, y int32, elementID *int32, arg1 *int32, arg2 *int32)  {
	retVal, _ := this.Call(0x00000581, []interface{}{x, y, elementID, arg1, arg2})
	_= retVal
}

var Chart_SetSourceData_OptArgs= []string{
	"PlotBy", 
}

func (this *Chart) SetSourceData(source string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_SetSourceData_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000585, []interface{}{source}, optArgs...)
	_= retVal
}

func (this *Chart) PlotBy() int32 {
	retVal, _ := this.PropGet(0x000000ca, nil)
	return retVal.LValVal()
}

func (this *Chart) SetPlotBy(rhs int32)  {
	_ = this.PropPut(0x000000ca, []interface{}{rhs})
}

func (this *Chart) HasLegend() bool {
	retVal, _ := this.PropGet(0x00000035, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetHasLegend(rhs bool)  {
	_ = this.PropPut(0x00000035, []interface{}{rhs})
}

func (this *Chart) Legend() *Legend {
	retVal, _ := this.PropGet(0x00000054, nil)
	return NewLegend(retVal.IDispatch(), false, true)
}

var Chart_Axes_OptArgs= []string{
	"Type", "AxisGroup", 
}

func (this *Chart) Axes(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_Axes_OptArgs, optArgs)
	retVal, _ := this.Call(0x60020035, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart_HasAxis_OptArgs= []string{
	"Index1", "Index2", 
}

func (this *Chart) HasAxis(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Chart_HasAxis_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x60020036, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Chart_SetHasAxis_OptArgs= []string{
	"Index1", "Index2", 
}

func (this *Chart) SetHasAxis(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_SetHasAxis_OptArgs, optArgs)
	_ = this.PropPut(0x60020036, nil, optArgs...)
}

func (this *Chart) Walls() *Walls {
	retVal, _ := this.PropGet(0x60020038, nil)
	return NewWalls(retVal.IDispatch(), false, true)
}

func (this *Chart) Floor() *Floor {
	retVal, _ := this.PropGet(0x60020039, nil)
	return NewFloor(retVal.IDispatch(), false, true)
}

func (this *Chart) PlotArea() *PlotArea {
	retVal, _ := this.PropGet(0x6002003a, nil)
	return NewPlotArea(retVal.IDispatch(), false, true)
}

func (this *Chart) PlotVisibleOnly() bool {
	retVal, _ := this.PropGet(0x0000005c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetPlotVisibleOnly(rhs bool)  {
	_ = this.PropPut(0x0000005c, []interface{}{rhs})
}

func (this *Chart) ChartArea() *ChartArea {
	retVal, _ := this.PropGet(0x6002003d, nil)
	return NewChartArea(retVal.IDispatch(), false, true)
}

var Chart_AutoFormat_OptArgs= []string{
	"Format", 
}

func (this *Chart) AutoFormat(gallery int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_AutoFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x6002003e, []interface{}{gallery}, optArgs...)
	_= retVal
}

func (this *Chart) AutoScaling() bool {
	retVal, _ := this.PropGet(0x6002003f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetAutoScaling(rhs bool)  {
	_ = this.PropPut(0x6002003f, []interface{}{rhs})
}

func (this *Chart) SetBackgroundPicture(fileName string)  {
	retVal, _ := this.Call(0x60020041, []interface{}{fileName})
	_= retVal
}

var Chart_ChartWizard_OptArgs= []string{
	"Source", "Gallery", "Format", "PlotBy", 
	"CategoryLabels", "SeriesLabels", "HasLegend", "Title", 
	"CategoryTitle", "ValueTitle", "ExtraTitle", 
}

func (this *Chart) ChartWizard(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_ChartWizard_OptArgs, optArgs)
	retVal, _ := this.Call(0x60020042, nil, optArgs...)
	_= retVal
}

var Chart_CopyPicture_OptArgs= []string{
	"Appearance", "Format", "Size", 
}

func (this *Chart) CopyPicture(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_CopyPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x60020043, nil, optArgs...)
	_= retVal
}

func (this *Chart) DataTable() *DataTable {
	retVal, _ := this.PropGet(0x60020044, nil)
	return NewDataTable(retVal.IDispatch(), false, true)
}

var Chart_Paste_OptArgs= []string{
	"Type", 
}

func (this *Chart) Paste(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_Paste_OptArgs, optArgs)
	retVal, _ := this.Call(0x60020047, nil, optArgs...)
	_= retVal
}

func (this *Chart) BarShape() int32 {
	retVal, _ := this.PropGet(0x60020048, nil)
	return retVal.LValVal()
}

func (this *Chart) SetBarShape(rhs int32)  {
	_ = this.PropPut(0x60020048, []interface{}{rhs})
}

var Chart_Export_OptArgs= []string{
	"FilterName", "Interactive", 
}

func (this *Chart) Export(fileName string, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Chart_Export_OptArgs, optArgs)
	retVal, _ := this.Call(0x6002004a, []interface{}{fileName}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetDefaultChart(name interface{})  {
	retVal, _ := this.Call(0x6002004b, []interface{}{name})
	_= retVal
}

func (this *Chart) ApplyChartTemplate(fileName string)  {
	retVal, _ := this.Call(0x6002004c, []interface{}{fileName})
	_= retVal
}

func (this *Chart) SaveChartTemplate(fileName string)  {
	retVal, _ := this.Call(0x6002004d, []interface{}{fileName})
	_= retVal
}

func (this *Chart) SideWall() *Walls {
	retVal, _ := this.PropGet(0x00000949, nil)
	return NewWalls(retVal.IDispatch(), false, true)
}

func (this *Chart) BackWall() *Walls {
	retVal, _ := this.PropGet(0x0000094a, nil)
	return NewWalls(retVal.IDispatch(), false, true)
}

func (this *Chart) ChartStyle() ole.Variant {
	retVal, _ := this.PropGet(0x000009a1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Chart) SetChartStyle(rhs interface{})  {
	_ = this.PropPut(0x000009a1, []interface{}{rhs})
}

func (this *Chart) ClearToMatchStyle()  {
	retVal, _ := this.Call(0x000009a2, nil)
	_= retVal
}

func (this *Chart) PivotLayout() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000716, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) HasPivotFields() bool {
	retVal, _ := this.PropGet(0x00000717, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetHasPivotFields(rhs bool)  {
	_ = this.PropPut(0x00000717, []interface{}{rhs})
}

func (this *Chart) ShowDataLabelsOverMaximum() bool {
	retVal, _ := this.PropGet(0x60020057, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetShowDataLabelsOverMaximum(rhs bool)  {
	_ = this.PropPut(0x60020057, []interface{}{rhs})
}

var Chart_ApplyLayout_OptArgs= []string{
	"ChartType", 
}

func (this *Chart) ApplyLayout(layout int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_ApplyLayout_OptArgs, optArgs)
	retVal, _ := this.Call(0x000009a4, []interface{}{layout}, optArgs...)
	_= retVal
}

func (this *Chart) Refresh()  {
	retVal, _ := this.Call(0x6002005b, nil)
	_= retVal
}

func (this *Chart) SetElement(element int32)  {
	retVal, _ := this.Call(0x6002005c, []interface{}{element})
	_= retVal
}

func (this *Chart) ChartData() *ChartData {
	retVal, _ := this.PropGet(0x6002005d, nil)
	return NewChartData(retVal.IDispatch(), false, true)
}

func (this *Chart) Shapes() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x6002005f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Chart) Area3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000011, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart_AreaGroups_OptArgs= []string{
	"Index", 
}

func (this *Chart) AreaGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_AreaGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000009, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) Bar3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000012, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart_BarGroups_OptArgs= []string{
	"Index", 
}

func (this *Chart) BarGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_BarGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000a, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) Column3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000013, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart_ColumnGroups_OptArgs= []string{
	"Index", 
}

func (this *Chart) ColumnGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_ColumnGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000b, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) Line3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000014, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart_LineGroups_OptArgs= []string{
	"Index", 
}

func (this *Chart) LineGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_LineGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000c, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) Pie3DGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000015, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart_PieGroups_OptArgs= []string{
	"Index", 
}

func (this *Chart) PieGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_PieGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000d, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart_DoughnutGroups_OptArgs= []string{
	"Index", 
}

func (this *Chart) DoughnutGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_DoughnutGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000e, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Chart_RadarGroups_OptArgs= []string{
	"Index", 
}

func (this *Chart) RadarGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_RadarGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000f, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) SurfaceGroup() *ChartGroup {
	retVal, _ := this.PropGet(0x00000016, nil)
	return NewChartGroup(retVal.IDispatch(), false, true)
}

var Chart_XYGroups_OptArgs= []string{
	"Index", 
}

func (this *Chart) XYGroups(optArgs ...interface{}) *ole.DispatchClass {
	optArgs = ole.ProcessOptArgs(Chart_XYGroups_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000010, nil, optArgs...)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Chart) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Chart_Copy_OptArgs= []string{
	"Before", "After", 
}

func (this *Chart) Copy(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Chart_Copy_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000227, nil, optArgs...)
	_= retVal
}

var Chart_Select_OptArgs= []string{
	"Replace", 
}

func (this *Chart) Select(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Chart_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000eb, nil, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Chart) ShowReportFilterFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b1c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetShowReportFilterFieldButtons(rhs bool)  {
	_ = this.PropPut(0x00000b1c, []interface{}{rhs})
}

func (this *Chart) ShowLegendFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b1d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetShowLegendFieldButtons(rhs bool)  {
	_ = this.PropPut(0x00000b1d, []interface{}{rhs})
}

func (this *Chart) ShowAxisFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b1e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetShowAxisFieldButtons(rhs bool)  {
	_ = this.PropPut(0x00000b1e, []interface{}{rhs})
}

func (this *Chart) ShowValueFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b1f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetShowValueFieldButtons(rhs bool)  {
	_ = this.PropPut(0x00000b1f, []interface{}{rhs})
}

func (this *Chart) ShowAllFieldButtons() bool {
	retVal, _ := this.PropGet(0x00000b20, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Chart) SetShowAllFieldButtons(rhs bool)  {
	_ = this.PropPut(0x00000b20, []interface{}{rhs})
}

