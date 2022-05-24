package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209A5-0000-0000-C000-000000000046
var IID_View = syscall.GUID{0x000209A5, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type View struct {
	ole.OleClient
}

func NewView(pDisp *win32.IDispatch, addRef bool, scoped bool) *View {
	 if pDisp == nil {
		return nil;
	}
	p := &View{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ViewFromVar(v ole.Variant) *View {
	return NewView(v.IDispatch(), false, false)
}

func (this *View) IID() *syscall.GUID {
	return &IID_View
}

func (this *View) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *View) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *View) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *View) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *View) Type() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *View) SetType(rhs int32)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *View) FullScreen() bool {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetFullScreen(rhs bool)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *View) Draft() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetDraft(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *View) ShowAll() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowAll(rhs bool)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *View) ShowFieldCodes() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowFieldCodes(rhs bool)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *View) MailMergeDataView() bool {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetMailMergeDataView(rhs bool)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *View) Magnifier() bool {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetMagnifier(rhs bool)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *View) ShowFirstLineOnly() bool {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowFirstLineOnly(rhs bool)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *View) ShowFormat() bool {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowFormat(rhs bool)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *View) Zoom() *Zoom {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewZoom(retVal.IDispatch(), false, true)
}

func (this *View) ShowObjectAnchors() bool {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowObjectAnchors(rhs bool)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *View) ShowTextBoundaries() bool {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowTextBoundaries(rhs bool)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *View) ShowHighlight() bool {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowHighlight(rhs bool)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *View) ShowDrawings() bool {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowDrawings(rhs bool)  {
	_ = this.PropPut(0x0000000e, []interface{}{rhs})
}

func (this *View) ShowTabs() bool {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowTabs(rhs bool)  {
	_ = this.PropPut(0x0000000f, []interface{}{rhs})
}

func (this *View) ShowSpaces() bool {
	retVal, _ := this.PropGet(0x00000010, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowSpaces(rhs bool)  {
	_ = this.PropPut(0x00000010, []interface{}{rhs})
}

func (this *View) ShowParagraphs() bool {
	retVal, _ := this.PropGet(0x00000011, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowParagraphs(rhs bool)  {
	_ = this.PropPut(0x00000011, []interface{}{rhs})
}

func (this *View) ShowHyphens() bool {
	retVal, _ := this.PropGet(0x00000012, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowHyphens(rhs bool)  {
	_ = this.PropPut(0x00000012, []interface{}{rhs})
}

func (this *View) ShowHiddenText() bool {
	retVal, _ := this.PropGet(0x00000013, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowHiddenText(rhs bool)  {
	_ = this.PropPut(0x00000013, []interface{}{rhs})
}

func (this *View) WrapToWindow() bool {
	retVal, _ := this.PropGet(0x00000014, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetWrapToWindow(rhs bool)  {
	_ = this.PropPut(0x00000014, []interface{}{rhs})
}

func (this *View) ShowPicturePlaceHolders() bool {
	retVal, _ := this.PropGet(0x00000015, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowPicturePlaceHolders(rhs bool)  {
	_ = this.PropPut(0x00000015, []interface{}{rhs})
}

func (this *View) ShowBookmarks() bool {
	retVal, _ := this.PropGet(0x00000016, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowBookmarks(rhs bool)  {
	_ = this.PropPut(0x00000016, []interface{}{rhs})
}

func (this *View) FieldShading() int32 {
	retVal, _ := this.PropGet(0x00000017, nil)
	return retVal.LValVal()
}

func (this *View) SetFieldShading(rhs int32)  {
	_ = this.PropPut(0x00000017, []interface{}{rhs})
}

func (this *View) ShowAnimation() bool {
	retVal, _ := this.PropGet(0x00000018, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowAnimation(rhs bool)  {
	_ = this.PropPut(0x00000018, []interface{}{rhs})
}

func (this *View) TableGridlines() bool {
	retVal, _ := this.PropGet(0x00000019, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetTableGridlines(rhs bool)  {
	_ = this.PropPut(0x00000019, []interface{}{rhs})
}

func (this *View) EnlargeFontsLessThan() int32 {
	retVal, _ := this.PropGet(0x0000001a, nil)
	return retVal.LValVal()
}

func (this *View) SetEnlargeFontsLessThan(rhs int32)  {
	_ = this.PropPut(0x0000001a, []interface{}{rhs})
}

func (this *View) ShowMainTextLayer() bool {
	retVal, _ := this.PropGet(0x0000001b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowMainTextLayer(rhs bool)  {
	_ = this.PropPut(0x0000001b, []interface{}{rhs})
}

func (this *View) SeekView() int32 {
	retVal, _ := this.PropGet(0x0000001c, nil)
	return retVal.LValVal()
}

func (this *View) SetSeekView(rhs int32)  {
	_ = this.PropPut(0x0000001c, []interface{}{rhs})
}

func (this *View) SplitSpecial() int32 {
	retVal, _ := this.PropGet(0x0000001d, nil)
	return retVal.LValVal()
}

func (this *View) SetSplitSpecial(rhs int32)  {
	_ = this.PropPut(0x0000001d, []interface{}{rhs})
}

func (this *View) BrowseToWindow() int32 {
	retVal, _ := this.PropGet(0x0000001e, nil)
	return retVal.LValVal()
}

func (this *View) SetBrowseToWindow(rhs int32)  {
	_ = this.PropPut(0x0000001e, []interface{}{rhs})
}

func (this *View) ShowOptionalBreaks() bool {
	retVal, _ := this.PropGet(0x0000001f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowOptionalBreaks(rhs bool)  {
	_ = this.PropPut(0x0000001f, []interface{}{rhs})
}

var View_CollapseOutline_OptArgs= []string{
	"Range", 
}

func (this *View) CollapseOutline(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(View_CollapseOutline_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, nil, optArgs...)
	_= retVal
}

var View_ExpandOutline_OptArgs= []string{
	"Range", 
}

func (this *View) ExpandOutline(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(View_ExpandOutline_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, nil, optArgs...)
	_= retVal
}

func (this *View) ShowAllHeadings()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

func (this *View) ShowHeading(level int32)  {
	retVal, _ := this.Call(0x00000068, []interface{}{level})
	_= retVal
}

func (this *View) PreviousHeaderFooter()  {
	retVal, _ := this.Call(0x00000069, nil)
	_= retVal
}

func (this *View) NextHeaderFooter()  {
	retVal, _ := this.Call(0x0000006a, nil)
	_= retVal
}

func (this *View) DisplayPageBoundaries() bool {
	retVal, _ := this.PropGet(0x00000020, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetDisplayPageBoundaries(rhs bool)  {
	_ = this.PropPut(0x00000020, []interface{}{rhs})
}

func (this *View) DisplaySmartTags() bool {
	retVal, _ := this.PropGet(0x00000021, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetDisplaySmartTags(rhs bool)  {
	_ = this.PropPut(0x00000021, []interface{}{rhs})
}

func (this *View) ShowRevisionsAndComments() bool {
	retVal, _ := this.PropGet(0x00000022, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowRevisionsAndComments(rhs bool)  {
	_ = this.PropPut(0x00000022, []interface{}{rhs})
}

func (this *View) ShowComments() bool {
	retVal, _ := this.PropGet(0x00000023, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowComments(rhs bool)  {
	_ = this.PropPut(0x00000023, []interface{}{rhs})
}

func (this *View) ShowInsertionsAndDeletions() bool {
	retVal, _ := this.PropGet(0x00000024, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowInsertionsAndDeletions(rhs bool)  {
	_ = this.PropPut(0x00000024, []interface{}{rhs})
}

func (this *View) ShowFormatChanges() bool {
	retVal, _ := this.PropGet(0x00000025, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowFormatChanges(rhs bool)  {
	_ = this.PropPut(0x00000025, []interface{}{rhs})
}

func (this *View) RevisionsView() int32 {
	retVal, _ := this.PropGet(0x00000026, nil)
	return retVal.LValVal()
}

func (this *View) SetRevisionsView(rhs int32)  {
	_ = this.PropPut(0x00000026, []interface{}{rhs})
}

func (this *View) RevisionsMode() int32 {
	retVal, _ := this.PropGet(0x00000027, nil)
	return retVal.LValVal()
}

func (this *View) SetRevisionsMode(rhs int32)  {
	_ = this.PropPut(0x00000027, []interface{}{rhs})
}

func (this *View) RevisionsBalloonWidth() float32 {
	retVal, _ := this.PropGet(0x00000028, nil)
	return retVal.FltValVal()
}

func (this *View) SetRevisionsBalloonWidth(rhs float32)  {
	_ = this.PropPut(0x00000028, []interface{}{rhs})
}

func (this *View) RevisionsBalloonWidthType() int32 {
	retVal, _ := this.PropGet(0x00000029, nil)
	return retVal.LValVal()
}

func (this *View) SetRevisionsBalloonWidthType(rhs int32)  {
	_ = this.PropPut(0x00000029, []interface{}{rhs})
}

func (this *View) RevisionsBalloonSide() int32 {
	retVal, _ := this.PropGet(0x0000002a, nil)
	return retVal.LValVal()
}

func (this *View) SetRevisionsBalloonSide(rhs int32)  {
	_ = this.PropPut(0x0000002a, []interface{}{rhs})
}

func (this *View) Reviewers() *Reviewers {
	retVal, _ := this.PropGet(0x0000002b, nil)
	return NewReviewers(retVal.IDispatch(), false, true)
}

func (this *View) RevisionsBalloonShowConnectingLines() bool {
	retVal, _ := this.PropGet(0x0000002c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetRevisionsBalloonShowConnectingLines(rhs bool)  {
	_ = this.PropPut(0x0000002c, []interface{}{rhs})
}

func (this *View) ReadingLayout() bool {
	retVal, _ := this.PropGet(0x0000002d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetReadingLayout(rhs bool)  {
	_ = this.PropPut(0x0000002d, []interface{}{rhs})
}

func (this *View) ShowXMLMarkup() int32 {
	retVal, _ := this.PropGet(0x0000002e, nil)
	return retVal.LValVal()
}

func (this *View) SetShowXMLMarkup(rhs int32)  {
	_ = this.PropPut(0x0000002e, []interface{}{rhs})
}

func (this *View) ShadeEditableRanges() int32 {
	retVal, _ := this.PropGet(0x0000002f, nil)
	return retVal.LValVal()
}

func (this *View) SetShadeEditableRanges(rhs int32)  {
	_ = this.PropPut(0x0000002f, []interface{}{rhs})
}

func (this *View) ShowInkAnnotations() bool {
	retVal, _ := this.PropGet(0x00000030, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowInkAnnotations(rhs bool)  {
	_ = this.PropPut(0x00000030, []interface{}{rhs})
}

func (this *View) DisplayBackgrounds() bool {
	retVal, _ := this.PropGet(0x00000031, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetDisplayBackgrounds(rhs bool)  {
	_ = this.PropPut(0x00000031, []interface{}{rhs})
}

func (this *View) ReadingLayoutActualView() bool {
	retVal, _ := this.PropGet(0x00000032, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetReadingLayoutActualView(rhs bool)  {
	_ = this.PropPut(0x00000032, []interface{}{rhs})
}

func (this *View) ReadingLayoutAllowMultiplePages() bool {
	retVal, _ := this.PropGet(0x00000033, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetReadingLayoutAllowMultiplePages(rhs bool)  {
	_ = this.PropPut(0x00000033, []interface{}{rhs})
}

func (this *View) ReadingLayoutAllowEditing() bool {
	retVal, _ := this.PropGet(0x00000035, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetReadingLayoutAllowEditing(rhs bool)  {
	_ = this.PropPut(0x00000035, []interface{}{rhs})
}

func (this *View) ReadingLayoutTruncateMargins() int32 {
	retVal, _ := this.PropGet(0x00000036, nil)
	return retVal.LValVal()
}

func (this *View) SetReadingLayoutTruncateMargins(rhs int32)  {
	_ = this.PropPut(0x00000036, []interface{}{rhs})
}

func (this *View) ShowMarkupAreaHighlight() bool {
	retVal, _ := this.PropGet(0x00000034, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowMarkupAreaHighlight(rhs bool)  {
	_ = this.PropPut(0x00000034, []interface{}{rhs})
}

func (this *View) Panning() bool {
	retVal, _ := this.PropGet(0x00000037, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetPanning(rhs bool)  {
	_ = this.PropPut(0x00000037, []interface{}{rhs})
}

func (this *View) ShowCropMarks() bool {
	retVal, _ := this.PropGet(0x00000038, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowCropMarks(rhs bool)  {
	_ = this.PropPut(0x00000038, []interface{}{rhs})
}

func (this *View) MarkupMode() int32 {
	retVal, _ := this.PropGet(0x00000039, nil)
	return retVal.LValVal()
}

func (this *View) SetMarkupMode(rhs int32)  {
	_ = this.PropPut(0x00000039, []interface{}{rhs})
}

func (this *View) ConflictMode() bool {
	retVal, _ := this.PropGet(0x0000003a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetConflictMode(rhs bool)  {
	_ = this.PropPut(0x0000003a, []interface{}{rhs})
}

func (this *View) ShowOtherAuthors() bool {
	retVal, _ := this.PropGet(0x0000003b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *View) SetShowOtherAuthors(rhs bool)  {
	_ = this.PropPut(0x0000003b, []interface{}{rhs})
}

