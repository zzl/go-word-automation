package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020971-0000-0000-C000-000000000046
var IID_PageSetup = syscall.GUID{0x00020971, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PageSetup struct {
	ole.OleClient
}

func NewPageSetup(pDisp *win32.IDispatch, addRef bool, scoped bool) *PageSetup {
	p := &PageSetup{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PageSetupFromVar(v ole.Variant) *PageSetup {
	return NewPageSetup(v.PdispValVal(), false, false)
}

func (this *PageSetup) IID() *syscall.GUID {
	return &IID_PageSetup
}

func (this *PageSetup) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PageSetup) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *PageSetup) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *PageSetup) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PageSetup) TopMargin() float32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetTopMargin(rhs float32)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) BottomMargin() float32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetBottomMargin(rhs float32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) LeftMargin() float32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetLeftMargin(rhs float32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) RightMargin() float32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetRightMargin(rhs float32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) Gutter() float32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetGutter(rhs float32)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) PageWidth() float32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetPageWidth(rhs float32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) PageHeight() float32 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetPageHeight(rhs float32)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) Orientation() int32 {
	retVal := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetOrientation(rhs int32)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) FirstPageTray() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetFirstPageTray(rhs int32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) OtherPagesTray() int32 {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetOtherPagesTray(rhs int32)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) VerticalAlignment() int32 {
	retVal := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetVerticalAlignment(rhs int32)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) MirrorMargins() int32 {
	retVal := this.PropGet(0x0000006f, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetMirrorMargins(rhs int32)  {
	retVal := this.PropPut(0x0000006f, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) HeaderDistance() float32 {
	retVal := this.PropGet(0x00000070, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetHeaderDistance(rhs float32)  {
	retVal := this.PropPut(0x00000070, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) FooterDistance() float32 {
	retVal := this.PropGet(0x00000071, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetFooterDistance(rhs float32)  {
	retVal := this.PropPut(0x00000071, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) SectionStart() int32 {
	retVal := this.PropGet(0x00000072, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetSectionStart(rhs int32)  {
	retVal := this.PropPut(0x00000072, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) OddAndEvenPagesHeaderFooter() int32 {
	retVal := this.PropGet(0x00000073, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetOddAndEvenPagesHeaderFooter(rhs int32)  {
	retVal := this.PropPut(0x00000073, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) DifferentFirstPageHeaderFooter() int32 {
	retVal := this.PropGet(0x00000074, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetDifferentFirstPageHeaderFooter(rhs int32)  {
	retVal := this.PropPut(0x00000074, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) SuppressEndnotes() int32 {
	retVal := this.PropGet(0x00000075, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetSuppressEndnotes(rhs int32)  {
	retVal := this.PropPut(0x00000075, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) LineNumbering() *LineNumbering {
	retVal := this.PropGet(0x00000076, nil)
	return NewLineNumbering(retVal.PdispValVal(), false, true)
}

func (this *PageSetup) SetLineNumbering(rhs *LineNumbering)  {
	retVal := this.PropPut(0x00000076, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) TextColumns() *TextColumns {
	retVal := this.PropGet(0x00000077, nil)
	return NewTextColumns(retVal.PdispValVal(), false, true)
}

func (this *PageSetup) SetTextColumns(rhs *TextColumns)  {
	retVal := this.PropPut(0x00000077, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) PaperSize() int32 {
	retVal := this.PropGet(0x00000078, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetPaperSize(rhs int32)  {
	retVal := this.PropPut(0x00000078, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) TwoPagesOnOne() bool {
	retVal := this.PropGet(0x00000079, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetTwoPagesOnOne(rhs bool)  {
	retVal := this.PropPut(0x00000079, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) GutterOnTop() bool {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetGutterOnTop(rhs bool)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) CharsLine() float32 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetCharsLine(rhs float32)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) LinesPage() float32 {
	retVal := this.PropGet(0x0000007c, nil)
	return retVal.FltValVal()
}

func (this *PageSetup) SetLinesPage(rhs float32)  {
	retVal := this.PropPut(0x0000007c, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) ShowGrid() bool {
	retVal := this.PropGet(0x00000080, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetShowGrid(rhs bool)  {
	retVal := this.PropPut(0x00000080, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) TogglePortrait()  {
	retVal := this.Call(0x000000c9, nil)
	_= retVal
}

func (this *PageSetup) SetAsTemplateDefault()  {
	retVal := this.Call(0x000000ca, nil)
	_= retVal
}

func (this *PageSetup) GutterStyle() int32 {
	retVal := this.PropGet(0x00000081, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetGutterStyle(rhs int32)  {
	retVal := this.PropPut(0x00000081, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) SectionDirection() int32 {
	retVal := this.PropGet(0x00000082, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetSectionDirection(rhs int32)  {
	retVal := this.PropPut(0x00000082, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) LayoutMode() int32 {
	retVal := this.PropGet(0x00000083, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetLayoutMode(rhs int32)  {
	retVal := this.PropPut(0x00000083, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) GutterPos() int32 {
	retVal := this.PropGet(0x000004c6, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetGutterPos(rhs int32)  {
	retVal := this.PropPut(0x000004c6, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) BookFoldPrinting() bool {
	retVal := this.PropGet(0x000004c7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetBookFoldPrinting(rhs bool)  {
	retVal := this.PropPut(0x000004c7, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) BookFoldRevPrinting() bool {
	retVal := this.PropGet(0x000004c8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageSetup) SetBookFoldRevPrinting(rhs bool)  {
	retVal := this.PropPut(0x000004c8, []interface{}{rhs})
	_= retVal
}

func (this *PageSetup) BookFoldPrintingSheets() int32 {
	retVal := this.PropGet(0x000004c9, nil)
	return retVal.LValVal()
}

func (this *PageSetup) SetBookFoldPrintingSheets(rhs int32)  {
	retVal := this.PropPut(0x000004c9, []interface{}{rhs})
	_= retVal
}

