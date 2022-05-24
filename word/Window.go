package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020962-0000-0000-C000-000000000046
var IID_Window = syscall.GUID{0x00020962, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Window struct {
	ole.OleClient
}

func NewWindow(pDisp *win32.IDispatch, addRef bool, scoped bool) *Window {
	 if pDisp == nil {
		return nil;
	}
	p := &Window{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WindowFromVar(v ole.Variant) *Window {
	return NewWindow(v.IDispatch(), false, false)
}

func (this *Window) IID() *syscall.GUID {
	return &IID_Window
}

func (this *Window) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Window) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Window) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Window) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Window) ActivePane() *Pane {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewPane(retVal.IDispatch(), false, true)
}

func (this *Window) Document() *Document {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewDocument(retVal.IDispatch(), false, true)
}

func (this *Window) Panes() *Panes {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewPanes(retVal.IDispatch(), false, true)
}

func (this *Window) Selection() *Selection {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewSelection(retVal.IDispatch(), false, true)
}

func (this *Window) Left() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Window) SetLeft(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *Window) Top() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Window) SetTop(rhs int32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *Window) Width() int32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *Window) SetWidth(rhs int32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *Window) Height() int32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Window) SetHeight(rhs int32)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *Window) Split() bool {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetSplit(rhs bool)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *Window) SplitVertical() int32 {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *Window) SetSplitVertical(rhs int32)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *Window) Caption() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Window) SetCaption(rhs string)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *Window) WindowState() int32 {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return retVal.LValVal()
}

func (this *Window) SetWindowState(rhs int32)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *Window) DisplayRulers() bool {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayRulers(rhs bool)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *Window) DisplayVerticalRuler() bool {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayVerticalRuler(rhs bool)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *Window) View() *View {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return NewView(retVal.IDispatch(), false, true)
}

func (this *Window) Type() int32 {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return retVal.LValVal()
}

func (this *Window) Next() *Window {
	retVal, _ := this.PropGet(0x00000010, nil)
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Window) Previous() *Window {
	retVal, _ := this.PropGet(0x00000011, nil)
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Window) WindowNumber() int32 {
	retVal, _ := this.PropGet(0x00000012, nil)
	return retVal.LValVal()
}

func (this *Window) DisplayVerticalScrollBar() bool {
	retVal, _ := this.PropGet(0x00000013, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayVerticalScrollBar(rhs bool)  {
	_ = this.PropPut(0x00000013, []interface{}{rhs})
}

func (this *Window) DisplayHorizontalScrollBar() bool {
	retVal, _ := this.PropGet(0x00000014, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayHorizontalScrollBar(rhs bool)  {
	_ = this.PropPut(0x00000014, []interface{}{rhs})
}

func (this *Window) StyleAreaWidth() float32 {
	retVal, _ := this.PropGet(0x00000015, nil)
	return retVal.FltValVal()
}

func (this *Window) SetStyleAreaWidth(rhs float32)  {
	_ = this.PropPut(0x00000015, []interface{}{rhs})
}

func (this *Window) DisplayScreenTips() bool {
	retVal, _ := this.PropGet(0x00000016, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayScreenTips(rhs bool)  {
	_ = this.PropPut(0x00000016, []interface{}{rhs})
}

func (this *Window) HorizontalPercentScrolled() int32 {
	retVal, _ := this.PropGet(0x00000017, nil)
	return retVal.LValVal()
}

func (this *Window) SetHorizontalPercentScrolled(rhs int32)  {
	_ = this.PropPut(0x00000017, []interface{}{rhs})
}

func (this *Window) VerticalPercentScrolled() int32 {
	retVal, _ := this.PropGet(0x00000018, nil)
	return retVal.LValVal()
}

func (this *Window) SetVerticalPercentScrolled(rhs int32)  {
	_ = this.PropPut(0x00000018, []interface{}{rhs})
}

func (this *Window) DocumentMap() bool {
	retVal, _ := this.PropGet(0x00000019, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDocumentMap(rhs bool)  {
	_ = this.PropPut(0x00000019, []interface{}{rhs})
}

func (this *Window) Active() bool {
	retVal, _ := this.PropGet(0x0000001a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) DocumentMapPercentWidth() int32 {
	retVal, _ := this.PropGet(0x0000001b, nil)
	return retVal.LValVal()
}

func (this *Window) SetDocumentMapPercentWidth(rhs int32)  {
	_ = this.PropPut(0x0000001b, []interface{}{rhs})
}

func (this *Window) Index() int32 {
	retVal, _ := this.PropGet(0x0000001c, nil)
	return retVal.LValVal()
}

func (this *Window) IMEMode() int32 {
	retVal, _ := this.PropGet(0x0000001e, nil)
	return retVal.LValVal()
}

func (this *Window) SetIMEMode(rhs int32)  {
	_ = this.PropPut(0x0000001e, []interface{}{rhs})
}

func (this *Window) Activate()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

var Window_Close_OptArgs= []string{
	"SaveChanges", "RouteDocument", 
}

func (this *Window) Close(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Window_Close_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, nil, optArgs...)
	_= retVal
}

var Window_LargeScroll_OptArgs= []string{
	"Down", "Up", "ToRight", "ToLeft", 
}

func (this *Window) LargeScroll(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Window_LargeScroll_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000067, nil, optArgs...)
	_= retVal
}

var Window_SmallScroll_OptArgs= []string{
	"Down", "Up", "ToRight", "ToLeft", 
}

func (this *Window) SmallScroll(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Window_SmallScroll_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000068, nil, optArgs...)
	_= retVal
}

func (this *Window) NewWindow() *Window {
	retVal, _ := this.Call(0x00000069, nil)
	return NewWindow(retVal.IDispatch(), false, true)
}

var Window_PrintOutOld_OptArgs= []string{
	"Background", "Append", "Range", "OutputFileName", 
	"From", "To", "Item", "Copies", 
	"Pages", "PageType", "PrintToFile", "Collate", 
	"ActivePrinterMacGX", "ManualDuplexPrint", 
}

func (this *Window) PrintOutOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Window_PrintOutOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006b, nil, optArgs...)
	_= retVal
}

var Window_PageScroll_OptArgs= []string{
	"Down", "Up", 
}

func (this *Window) PageScroll(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Window_PageScroll_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006c, nil, optArgs...)
	_= retVal
}

func (this *Window) SetFocus()  {
	retVal, _ := this.Call(0x0000006d, nil)
	_= retVal
}

func (this *Window) RangeFromPoint(x int32, y int32) *ole.DispatchClass {
	retVal, _ := this.Call(0x0000006e, []interface{}{x, y})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Window_ScrollIntoView_OptArgs= []string{
	"Start", 
}

func (this *Window) ScrollIntoView(obj *win32.IUnknown, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Window_ScrollIntoView_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006f, []interface{}{obj}, optArgs...)
	_= retVal
}

func (this *Window) GetPoint(screenPixelsLeft *int32, screenPixelsTop *int32, screenPixelsWidth *int32, screenPixelsHeight *int32, obj *win32.IUnknown)  {
	retVal, _ := this.Call(0x00000070, []interface{}{screenPixelsLeft, screenPixelsTop, screenPixelsWidth, screenPixelsHeight, obj})
	_= retVal
}

var Window_PrintOut2000_OptArgs= []string{
	"Background", "Append", "Range", "OutputFileName", 
	"From", "To", "Item", "Copies", 
	"Pages", "PageType", "PrintToFile", "Collate", 
	"ActivePrinterMacGX", "ManualDuplexPrint", "PrintZoomColumn", "PrintZoomRow", 
	"PrintZoomPaperWidth", "PrintZoomPaperHeight", 
}

func (this *Window) PrintOut2000(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Window_PrintOut2000_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bc, nil, optArgs...)
	_= retVal
}

func (this *Window) UsableWidth() int32 {
	retVal, _ := this.PropGet(0x0000001f, nil)
	return retVal.LValVal()
}

func (this *Window) UsableHeight() int32 {
	retVal, _ := this.PropGet(0x00000020, nil)
	return retVal.LValVal()
}

func (this *Window) EnvelopeVisible() bool {
	retVal, _ := this.PropGet(0x00000021, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetEnvelopeVisible(rhs bool)  {
	_ = this.PropPut(0x00000021, []interface{}{rhs})
}

func (this *Window) DisplayRightRuler() bool {
	retVal, _ := this.PropGet(0x00000023, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayRightRuler(rhs bool)  {
	_ = this.PropPut(0x00000023, []interface{}{rhs})
}

func (this *Window) DisplayLeftScrollBar() bool {
	retVal, _ := this.PropGet(0x00000022, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetDisplayLeftScrollBar(rhs bool)  {
	_ = this.PropPut(0x00000022, []interface{}{rhs})
}

func (this *Window) Visible() bool {
	retVal, _ := this.PropGet(0x00000024, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetVisible(rhs bool)  {
	_ = this.PropPut(0x00000024, []interface{}{rhs})
}

var Window_PrintOut_OptArgs= []string{
	"Background", "Append", "Range", "OutputFileName", 
	"From", "To", "Item", "Copies", 
	"Pages", "PageType", "PrintToFile", "Collate", 
	"ActivePrinterMacGX", "ManualDuplexPrint", "PrintZoomColumn", "PrintZoomRow", 
	"PrintZoomPaperWidth", "PrintZoomPaperHeight", 
}

func (this *Window) PrintOut(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Window_PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bd, nil, optArgs...)
	_= retVal
}

func (this *Window) ToggleShowAllReviewers()  {
	retVal, _ := this.Call(0x000001be, nil)
	_= retVal
}

func (this *Window) Thumbnails() bool {
	retVal, _ := this.PropGet(0x00000025, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Window) SetThumbnails(rhs bool)  {
	_ = this.PropPut(0x00000025, []interface{}{rhs})
}

func (this *Window) ShowSourceDocuments() int32 {
	retVal, _ := this.PropGet(0x00000026, nil)
	return retVal.LValVal()
}

func (this *Window) SetShowSourceDocuments(rhs int32)  {
	_ = this.PropPut(0x00000026, []interface{}{rhs})
}

func (this *Window) ToggleRibbon()  {
	retVal, _ := this.Call(0x000001bf, nil)
	_= retVal
}

