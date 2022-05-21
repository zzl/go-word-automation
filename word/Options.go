package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209B7-0000-0000-C000-000000000046
var IID_Options = syscall.GUID{0x000209B7, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Options struct {
	ole.OleClient
}

func NewOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *Options {
	p := &Options{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OptionsFromVar(v ole.Variant) *Options {
	return NewOptions(v.PdispValVal(), false, false)
}

func (this *Options) IID() *syscall.GUID {
	return &IID_Options
}

func (this *Options) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Options) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Options) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Options) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Options) AllowAccentedUppercase() bool {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAllowAccentedUppercase(rhs bool)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *Options) WPHelp() bool {
	retVal := this.PropGet(0x00000011, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetWPHelp(rhs bool)  {
	retVal := this.PropPut(0x00000011, []interface{}{rhs})
	_= retVal
}

func (this *Options) WPDocNavKeys() bool {
	retVal := this.PropGet(0x00000012, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetWPDocNavKeys(rhs bool)  {
	retVal := this.PropPut(0x00000012, []interface{}{rhs})
	_= retVal
}

func (this *Options) Pagination() bool {
	retVal := this.PropGet(0x00000013, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPagination(rhs bool)  {
	retVal := this.PropPut(0x00000013, []interface{}{rhs})
	_= retVal
}

func (this *Options) BlueScreen() bool {
	retVal := this.PropGet(0x00000014, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetBlueScreen(rhs bool)  {
	retVal := this.PropPut(0x00000014, []interface{}{rhs})
	_= retVal
}

func (this *Options) EnableSound() bool {
	retVal := this.PropGet(0x00000015, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetEnableSound(rhs bool)  {
	retVal := this.PropPut(0x00000015, []interface{}{rhs})
	_= retVal
}

func (this *Options) ConfirmConversions() bool {
	retVal := this.PropGet(0x00000016, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetConfirmConversions(rhs bool)  {
	retVal := this.PropPut(0x00000016, []interface{}{rhs})
	_= retVal
}

func (this *Options) UpdateLinksAtOpen() bool {
	retVal := this.PropGet(0x00000017, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetUpdateLinksAtOpen(rhs bool)  {
	retVal := this.PropPut(0x00000017, []interface{}{rhs})
	_= retVal
}

func (this *Options) SendMailAttach() bool {
	retVal := this.PropGet(0x00000018, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSendMailAttach(rhs bool)  {
	retVal := this.PropPut(0x00000018, []interface{}{rhs})
	_= retVal
}

func (this *Options) MeasurementUnit() int32 {
	retVal := this.PropGet(0x0000001a, nil)
	return retVal.LValVal()
}

func (this *Options) SetMeasurementUnit(rhs int32)  {
	retVal := this.PropPut(0x0000001a, []interface{}{rhs})
	_= retVal
}

func (this *Options) ButtonFieldClicks() int32 {
	retVal := this.PropGet(0x0000001b, nil)
	return retVal.LValVal()
}

func (this *Options) SetButtonFieldClicks(rhs int32)  {
	retVal := this.PropPut(0x0000001b, []interface{}{rhs})
	_= retVal
}

func (this *Options) ShortMenuNames() bool {
	retVal := this.PropGet(0x0000001c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetShortMenuNames(rhs bool)  {
	retVal := this.PropPut(0x0000001c, []interface{}{rhs})
	_= retVal
}

func (this *Options) RTFInClipboard() bool {
	retVal := this.PropGet(0x0000001d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetRTFInClipboard(rhs bool)  {
	retVal := this.PropPut(0x0000001d, []interface{}{rhs})
	_= retVal
}

func (this *Options) UpdateFieldsAtPrint() bool {
	retVal := this.PropGet(0x0000001e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetUpdateFieldsAtPrint(rhs bool)  {
	retVal := this.PropPut(0x0000001e, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintProperties() bool {
	retVal := this.PropGet(0x0000001f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintProperties(rhs bool)  {
	retVal := this.PropPut(0x0000001f, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintFieldCodes() bool {
	retVal := this.PropGet(0x00000020, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintFieldCodes(rhs bool)  {
	retVal := this.PropPut(0x00000020, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintComments() bool {
	retVal := this.PropGet(0x00000021, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintComments(rhs bool)  {
	retVal := this.PropPut(0x00000021, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintHiddenText() bool {
	retVal := this.PropGet(0x00000022, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintHiddenText(rhs bool)  {
	retVal := this.PropPut(0x00000022, []interface{}{rhs})
	_= retVal
}

func (this *Options) EnvelopeFeederInstalled() bool {
	retVal := this.PropGet(0x00000023, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) UpdateLinksAtPrint() bool {
	retVal := this.PropGet(0x00000024, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetUpdateLinksAtPrint(rhs bool)  {
	retVal := this.PropPut(0x00000024, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintBackground() bool {
	retVal := this.PropGet(0x00000025, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintBackground(rhs bool)  {
	retVal := this.PropPut(0x00000025, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintDrawingObjects() bool {
	retVal := this.PropGet(0x00000026, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintDrawingObjects(rhs bool)  {
	retVal := this.PropPut(0x00000026, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultTray() string {
	retVal := this.PropGet(0x00000027, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Options) SetDefaultTray(rhs string)  {
	retVal := this.PropPut(0x00000027, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultTrayID() int32 {
	retVal := this.PropGet(0x00000028, nil)
	return retVal.LValVal()
}

func (this *Options) SetDefaultTrayID(rhs int32)  {
	retVal := this.PropPut(0x00000028, []interface{}{rhs})
	_= retVal
}

func (this *Options) CreateBackup() bool {
	retVal := this.PropGet(0x00000029, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetCreateBackup(rhs bool)  {
	retVal := this.PropPut(0x00000029, []interface{}{rhs})
	_= retVal
}

func (this *Options) AllowFastSave() bool {
	retVal := this.PropGet(0x0000002a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAllowFastSave(rhs bool)  {
	retVal := this.PropPut(0x0000002a, []interface{}{rhs})
	_= retVal
}

func (this *Options) SavePropertiesPrompt() bool {
	retVal := this.PropGet(0x0000002b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSavePropertiesPrompt(rhs bool)  {
	retVal := this.PropPut(0x0000002b, []interface{}{rhs})
	_= retVal
}

func (this *Options) SaveNormalPrompt() bool {
	retVal := this.PropGet(0x0000002c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSaveNormalPrompt(rhs bool)  {
	retVal := this.PropPut(0x0000002c, []interface{}{rhs})
	_= retVal
}

func (this *Options) SaveInterval() int32 {
	retVal := this.PropGet(0x0000002d, nil)
	return retVal.LValVal()
}

func (this *Options) SetSaveInterval(rhs int32)  {
	retVal := this.PropPut(0x0000002d, []interface{}{rhs})
	_= retVal
}

func (this *Options) BackgroundSave() bool {
	retVal := this.PropGet(0x0000002e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetBackgroundSave(rhs bool)  {
	retVal := this.PropPut(0x0000002e, []interface{}{rhs})
	_= retVal
}

func (this *Options) InsertedTextMark() int32 {
	retVal := this.PropGet(0x00000039, nil)
	return retVal.LValVal()
}

func (this *Options) SetInsertedTextMark(rhs int32)  {
	retVal := this.PropPut(0x00000039, []interface{}{rhs})
	_= retVal
}

func (this *Options) DeletedTextMark() int32 {
	retVal := this.PropGet(0x0000003a, nil)
	return retVal.LValVal()
}

func (this *Options) SetDeletedTextMark(rhs int32)  {
	retVal := this.PropPut(0x0000003a, []interface{}{rhs})
	_= retVal
}

func (this *Options) RevisedLinesMark() int32 {
	retVal := this.PropGet(0x0000003b, nil)
	return retVal.LValVal()
}

func (this *Options) SetRevisedLinesMark(rhs int32)  {
	retVal := this.PropPut(0x0000003b, []interface{}{rhs})
	_= retVal
}

func (this *Options) InsertedTextColor() int32 {
	retVal := this.PropGet(0x0000003c, nil)
	return retVal.LValVal()
}

func (this *Options) SetInsertedTextColor(rhs int32)  {
	retVal := this.PropPut(0x0000003c, []interface{}{rhs})
	_= retVal
}

func (this *Options) DeletedTextColor() int32 {
	retVal := this.PropGet(0x0000003d, nil)
	return retVal.LValVal()
}

func (this *Options) SetDeletedTextColor(rhs int32)  {
	retVal := this.PropPut(0x0000003d, []interface{}{rhs})
	_= retVal
}

func (this *Options) RevisedLinesColor() int32 {
	retVal := this.PropGet(0x0000003e, nil)
	return retVal.LValVal()
}

func (this *Options) SetRevisedLinesColor(rhs int32)  {
	retVal := this.PropPut(0x0000003e, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultFilePath(path int32) string {
	retVal := this.PropGet(0x00000041, []interface{}{path})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Options) SetDefaultFilePath(path int32, rhs string)  {
	retVal := this.PropPut(0x00000041, []interface{}{path, rhs})
	_= retVal
}

func (this *Options) Overtype() bool {
	retVal := this.PropGet(0x00000042, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetOvertype(rhs bool)  {
	retVal := this.PropPut(0x00000042, []interface{}{rhs})
	_= retVal
}

func (this *Options) ReplaceSelection() bool {
	retVal := this.PropGet(0x00000043, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetReplaceSelection(rhs bool)  {
	retVal := this.PropPut(0x00000043, []interface{}{rhs})
	_= retVal
}

func (this *Options) AllowDragAndDrop() bool {
	retVal := this.PropGet(0x00000044, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAllowDragAndDrop(rhs bool)  {
	retVal := this.PropPut(0x00000044, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoWordSelection() bool {
	retVal := this.PropGet(0x00000045, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoWordSelection(rhs bool)  {
	retVal := this.PropPut(0x00000045, []interface{}{rhs})
	_= retVal
}

func (this *Options) INSKeyForPaste() bool {
	retVal := this.PropGet(0x00000046, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetINSKeyForPaste(rhs bool)  {
	retVal := this.PropPut(0x00000046, []interface{}{rhs})
	_= retVal
}

func (this *Options) SmartCutPaste() bool {
	retVal := this.PropGet(0x00000047, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSmartCutPaste(rhs bool)  {
	retVal := this.PropPut(0x00000047, []interface{}{rhs})
	_= retVal
}

func (this *Options) TabIndentKey() bool {
	retVal := this.PropGet(0x00000048, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetTabIndentKey(rhs bool)  {
	retVal := this.PropPut(0x00000048, []interface{}{rhs})
	_= retVal
}

func (this *Options) PictureEditor() string {
	retVal := this.PropGet(0x00000049, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Options) SetPictureEditor(rhs string)  {
	retVal := this.PropPut(0x00000049, []interface{}{rhs})
	_= retVal
}

func (this *Options) AnimateScreenMovements() bool {
	retVal := this.PropGet(0x0000004a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAnimateScreenMovements(rhs bool)  {
	retVal := this.PropPut(0x0000004a, []interface{}{rhs})
	_= retVal
}

func (this *Options) VirusProtection() bool {
	retVal := this.PropGet(0x0000004b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetVirusProtection(rhs bool)  {
	retVal := this.PropPut(0x0000004b, []interface{}{rhs})
	_= retVal
}

func (this *Options) RevisedPropertiesMark() int32 {
	retVal := this.PropGet(0x0000004c, nil)
	return retVal.LValVal()
}

func (this *Options) SetRevisedPropertiesMark(rhs int32)  {
	retVal := this.PropPut(0x0000004c, []interface{}{rhs})
	_= retVal
}

func (this *Options) RevisedPropertiesColor() int32 {
	retVal := this.PropGet(0x0000004d, nil)
	return retVal.LValVal()
}

func (this *Options) SetRevisedPropertiesColor(rhs int32)  {
	retVal := this.PropPut(0x0000004d, []interface{}{rhs})
	_= retVal
}

func (this *Options) SnapToGrid() bool {
	retVal := this.PropGet(0x0000004f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSnapToGrid(rhs bool)  {
	retVal := this.PropPut(0x0000004f, []interface{}{rhs})
	_= retVal
}

func (this *Options) SnapToShapes() bool {
	retVal := this.PropGet(0x00000050, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSnapToShapes(rhs bool)  {
	retVal := this.PropPut(0x00000050, []interface{}{rhs})
	_= retVal
}

func (this *Options) GridDistanceHorizontal() float32 {
	retVal := this.PropGet(0x00000051, nil)
	return retVal.FltValVal()
}

func (this *Options) SetGridDistanceHorizontal(rhs float32)  {
	retVal := this.PropPut(0x00000051, []interface{}{rhs})
	_= retVal
}

func (this *Options) GridDistanceVertical() float32 {
	retVal := this.PropGet(0x00000052, nil)
	return retVal.FltValVal()
}

func (this *Options) SetGridDistanceVertical(rhs float32)  {
	retVal := this.PropPut(0x00000052, []interface{}{rhs})
	_= retVal
}

func (this *Options) GridOriginHorizontal() float32 {
	retVal := this.PropGet(0x00000053, nil)
	return retVal.FltValVal()
}

func (this *Options) SetGridOriginHorizontal(rhs float32)  {
	retVal := this.PropPut(0x00000053, []interface{}{rhs})
	_= retVal
}

func (this *Options) GridOriginVertical() float32 {
	retVal := this.PropGet(0x00000054, nil)
	return retVal.FltValVal()
}

func (this *Options) SetGridOriginVertical(rhs float32)  {
	retVal := this.PropPut(0x00000054, []interface{}{rhs})
	_= retVal
}

func (this *Options) InlineConversion() bool {
	retVal := this.PropGet(0x00000056, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetInlineConversion(rhs bool)  {
	retVal := this.PropPut(0x00000056, []interface{}{rhs})
	_= retVal
}

func (this *Options) IMEAutomaticControl() bool {
	retVal := this.PropGet(0x00000057, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetIMEAutomaticControl(rhs bool)  {
	retVal := this.PropPut(0x00000057, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatApplyHeadings() bool {
	retVal := this.PropGet(0x000000fa, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatApplyHeadings(rhs bool)  {
	retVal := this.PropPut(0x000000fa, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatApplyLists() bool {
	retVal := this.PropGet(0x000000fb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatApplyLists(rhs bool)  {
	retVal := this.PropPut(0x000000fb, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatApplyBulletedLists() bool {
	retVal := this.PropGet(0x000000fc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatApplyBulletedLists(rhs bool)  {
	retVal := this.PropPut(0x000000fc, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatApplyOtherParas() bool {
	retVal := this.PropGet(0x000000fd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatApplyOtherParas(rhs bool)  {
	retVal := this.PropPut(0x000000fd, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatReplaceQuotes() bool {
	retVal := this.PropGet(0x000000fe, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatReplaceQuotes(rhs bool)  {
	retVal := this.PropPut(0x000000fe, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatReplaceSymbols() bool {
	retVal := this.PropGet(0x000000ff, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatReplaceSymbols(rhs bool)  {
	retVal := this.PropPut(0x000000ff, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatReplaceOrdinals() bool {
	retVal := this.PropGet(0x00000100, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatReplaceOrdinals(rhs bool)  {
	retVal := this.PropPut(0x00000100, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatReplaceFractions() bool {
	retVal := this.PropGet(0x00000101, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatReplaceFractions(rhs bool)  {
	retVal := this.PropPut(0x00000101, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatReplacePlainTextEmphasis() bool {
	retVal := this.PropGet(0x00000102, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatReplacePlainTextEmphasis(rhs bool)  {
	retVal := this.PropPut(0x00000102, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatPreserveStyles() bool {
	retVal := this.PropGet(0x00000103, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatPreserveStyles(rhs bool)  {
	retVal := this.PropPut(0x00000103, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeApplyHeadings() bool {
	retVal := this.PropGet(0x00000104, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeApplyHeadings(rhs bool)  {
	retVal := this.PropPut(0x00000104, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeApplyBorders() bool {
	retVal := this.PropGet(0x00000105, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeApplyBorders(rhs bool)  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeApplyBulletedLists() bool {
	retVal := this.PropGet(0x00000106, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeApplyBulletedLists(rhs bool)  {
	retVal := this.PropPut(0x00000106, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeApplyNumberedLists() bool {
	retVal := this.PropGet(0x00000107, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeApplyNumberedLists(rhs bool)  {
	retVal := this.PropPut(0x00000107, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeReplaceQuotes() bool {
	retVal := this.PropGet(0x00000108, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeReplaceQuotes(rhs bool)  {
	retVal := this.PropPut(0x00000108, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeReplaceSymbols() bool {
	retVal := this.PropGet(0x00000109, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeReplaceSymbols(rhs bool)  {
	retVal := this.PropPut(0x00000109, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeReplaceOrdinals() bool {
	retVal := this.PropGet(0x0000010a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeReplaceOrdinals(rhs bool)  {
	retVal := this.PropPut(0x0000010a, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeReplaceFractions() bool {
	retVal := this.PropGet(0x0000010b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeReplaceFractions(rhs bool)  {
	retVal := this.PropPut(0x0000010b, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeReplacePlainTextEmphasis() bool {
	retVal := this.PropGet(0x0000010c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeReplacePlainTextEmphasis(rhs bool)  {
	retVal := this.PropPut(0x0000010c, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeFormatListItemBeginning() bool {
	retVal := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeFormatListItemBeginning(rhs bool)  {
	retVal := this.PropPut(0x0000010d, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeDefineStyles() bool {
	retVal := this.PropGet(0x0000010e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeDefineStyles(rhs bool)  {
	retVal := this.PropPut(0x0000010e, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatPlainTextWordMail() bool {
	retVal := this.PropGet(0x0000010f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatPlainTextWordMail(rhs bool)  {
	retVal := this.PropPut(0x0000010f, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeReplaceHyperlinks() bool {
	retVal := this.PropGet(0x00000110, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeReplaceHyperlinks(rhs bool)  {
	retVal := this.PropPut(0x00000110, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatReplaceHyperlinks() bool {
	retVal := this.PropGet(0x00000111, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatReplaceHyperlinks(rhs bool)  {
	retVal := this.PropPut(0x00000111, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultHighlightColorIndex() int32 {
	retVal := this.PropGet(0x00000112, nil)
	return retVal.LValVal()
}

func (this *Options) SetDefaultHighlightColorIndex(rhs int32)  {
	retVal := this.PropPut(0x00000112, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultBorderLineStyle() int32 {
	retVal := this.PropGet(0x00000113, nil)
	return retVal.LValVal()
}

func (this *Options) SetDefaultBorderLineStyle(rhs int32)  {
	retVal := this.PropPut(0x00000113, []interface{}{rhs})
	_= retVal
}

func (this *Options) CheckSpellingAsYouType() bool {
	retVal := this.PropGet(0x00000114, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetCheckSpellingAsYouType(rhs bool)  {
	retVal := this.PropPut(0x00000114, []interface{}{rhs})
	_= retVal
}

func (this *Options) CheckGrammarAsYouType() bool {
	retVal := this.PropGet(0x00000115, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetCheckGrammarAsYouType(rhs bool)  {
	retVal := this.PropPut(0x00000115, []interface{}{rhs})
	_= retVal
}

func (this *Options) IgnoreInternetAndFileAddresses() bool {
	retVal := this.PropGet(0x00000116, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetIgnoreInternetAndFileAddresses(rhs bool)  {
	retVal := this.PropPut(0x00000116, []interface{}{rhs})
	_= retVal
}

func (this *Options) ShowReadabilityStatistics() bool {
	retVal := this.PropGet(0x00000117, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetShowReadabilityStatistics(rhs bool)  {
	retVal := this.PropPut(0x00000117, []interface{}{rhs})
	_= retVal
}

func (this *Options) IgnoreUppercase() bool {
	retVal := this.PropGet(0x00000118, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetIgnoreUppercase(rhs bool)  {
	retVal := this.PropPut(0x00000118, []interface{}{rhs})
	_= retVal
}

func (this *Options) IgnoreMixedDigits() bool {
	retVal := this.PropGet(0x00000119, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetIgnoreMixedDigits(rhs bool)  {
	retVal := this.PropPut(0x00000119, []interface{}{rhs})
	_= retVal
}

func (this *Options) SuggestFromMainDictionaryOnly() bool {
	retVal := this.PropGet(0x0000011a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSuggestFromMainDictionaryOnly(rhs bool)  {
	retVal := this.PropPut(0x0000011a, []interface{}{rhs})
	_= retVal
}

func (this *Options) SuggestSpellingCorrections() bool {
	retVal := this.PropGet(0x0000011b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSuggestSpellingCorrections(rhs bool)  {
	retVal := this.PropPut(0x0000011b, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultBorderLineWidth() int32 {
	retVal := this.PropGet(0x0000011c, nil)
	return retVal.LValVal()
}

func (this *Options) SetDefaultBorderLineWidth(rhs int32)  {
	retVal := this.PropPut(0x0000011c, []interface{}{rhs})
	_= retVal
}

func (this *Options) CheckGrammarWithSpelling() bool {
	retVal := this.PropGet(0x0000011d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetCheckGrammarWithSpelling(rhs bool)  {
	retVal := this.PropPut(0x0000011d, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultOpenFormat() int32 {
	retVal := this.PropGet(0x0000011e, nil)
	return retVal.LValVal()
}

func (this *Options) SetDefaultOpenFormat(rhs int32)  {
	retVal := this.PropPut(0x0000011e, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintDraft() bool {
	retVal := this.PropGet(0x0000011f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintDraft(rhs bool)  {
	retVal := this.PropPut(0x0000011f, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintReverse() bool {
	retVal := this.PropGet(0x00000120, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintReverse(rhs bool)  {
	retVal := this.PropPut(0x00000120, []interface{}{rhs})
	_= retVal
}

func (this *Options) MapPaperSize() bool {
	retVal := this.PropGet(0x00000121, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMapPaperSize(rhs bool)  {
	retVal := this.PropPut(0x00000121, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeApplyTables() bool {
	retVal := this.PropGet(0x00000122, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeApplyTables(rhs bool)  {
	retVal := this.PropPut(0x00000122, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatApplyFirstIndents() bool {
	retVal := this.PropGet(0x00000123, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatApplyFirstIndents(rhs bool)  {
	retVal := this.PropPut(0x00000123, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatMatchParentheses() bool {
	retVal := this.PropGet(0x00000126, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatMatchParentheses(rhs bool)  {
	retVal := this.PropPut(0x00000126, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatReplaceFarEastDashes() bool {
	retVal := this.PropGet(0x00000127, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatReplaceFarEastDashes(rhs bool)  {
	retVal := this.PropPut(0x00000127, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatDeleteAutoSpaces() bool {
	retVal := this.PropGet(0x00000128, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatDeleteAutoSpaces(rhs bool)  {
	retVal := this.PropPut(0x00000128, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeApplyFirstIndents() bool {
	retVal := this.PropGet(0x00000129, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeApplyFirstIndents(rhs bool)  {
	retVal := this.PropPut(0x00000129, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeApplyDates() bool {
	retVal := this.PropGet(0x0000012a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeApplyDates(rhs bool)  {
	retVal := this.PropPut(0x0000012a, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeApplyClosings() bool {
	retVal := this.PropGet(0x0000012b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeApplyClosings(rhs bool)  {
	retVal := this.PropPut(0x0000012b, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeMatchParentheses() bool {
	retVal := this.PropGet(0x0000012c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeMatchParentheses(rhs bool)  {
	retVal := this.PropPut(0x0000012c, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeReplaceFarEastDashes() bool {
	retVal := this.PropGet(0x0000012d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeReplaceFarEastDashes(rhs bool)  {
	retVal := this.PropPut(0x0000012d, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeDeleteAutoSpaces() bool {
	retVal := this.PropGet(0x0000012e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeDeleteAutoSpaces(rhs bool)  {
	retVal := this.PropPut(0x0000012e, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeInsertClosings() bool {
	retVal := this.PropGet(0x0000012f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeInsertClosings(rhs bool)  {
	retVal := this.PropPut(0x0000012f, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeAutoLetterWizard() bool {
	retVal := this.PropGet(0x00000130, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeAutoLetterWizard(rhs bool)  {
	retVal := this.PropPut(0x00000130, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoFormatAsYouTypeInsertOvers() bool {
	retVal := this.PropGet(0x00000131, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoFormatAsYouTypeInsertOvers(rhs bool)  {
	retVal := this.PropPut(0x00000131, []interface{}{rhs})
	_= retVal
}

func (this *Options) DisplayGridLines() bool {
	retVal := this.PropGet(0x00000132, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetDisplayGridLines(rhs bool)  {
	retVal := this.PropPut(0x00000132, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyCase() bool {
	retVal := this.PropGet(0x00000135, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyCase(rhs bool)  {
	retVal := this.PropPut(0x00000135, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyByte() bool {
	retVal := this.PropGet(0x00000136, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyByte(rhs bool)  {
	retVal := this.PropPut(0x00000136, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyHiragana() bool {
	retVal := this.PropGet(0x00000137, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyHiragana(rhs bool)  {
	retVal := this.PropPut(0x00000137, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzySmallKana() bool {
	retVal := this.PropGet(0x00000138, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzySmallKana(rhs bool)  {
	retVal := this.PropPut(0x00000138, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyDash() bool {
	retVal := this.PropGet(0x00000139, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyDash(rhs bool)  {
	retVal := this.PropPut(0x00000139, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyIterationMark() bool {
	retVal := this.PropGet(0x0000013a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyIterationMark(rhs bool)  {
	retVal := this.PropPut(0x0000013a, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyKanji() bool {
	retVal := this.PropGet(0x0000013b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyKanji(rhs bool)  {
	retVal := this.PropPut(0x0000013b, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyOldKana() bool {
	retVal := this.PropGet(0x0000013c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyOldKana(rhs bool)  {
	retVal := this.PropPut(0x0000013c, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyProlongedSoundMark() bool {
	retVal := this.PropGet(0x0000013d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyProlongedSoundMark(rhs bool)  {
	retVal := this.PropPut(0x0000013d, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyDZ() bool {
	retVal := this.PropGet(0x0000013e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyDZ(rhs bool)  {
	retVal := this.PropPut(0x0000013e, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyBV() bool {
	retVal := this.PropGet(0x0000013f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyBV(rhs bool)  {
	retVal := this.PropPut(0x0000013f, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyTC() bool {
	retVal := this.PropGet(0x00000140, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyTC(rhs bool)  {
	retVal := this.PropPut(0x00000140, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyHF() bool {
	retVal := this.PropGet(0x00000141, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyHF(rhs bool)  {
	retVal := this.PropPut(0x00000141, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyZJ() bool {
	retVal := this.PropGet(0x00000142, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyZJ(rhs bool)  {
	retVal := this.PropPut(0x00000142, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyAY() bool {
	retVal := this.PropGet(0x00000143, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyAY(rhs bool)  {
	retVal := this.PropPut(0x00000143, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyKiKu() bool {
	retVal := this.PropGet(0x00000144, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyKiKu(rhs bool)  {
	retVal := this.PropPut(0x00000144, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzyPunctuation() bool {
	retVal := this.PropGet(0x00000145, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzyPunctuation(rhs bool)  {
	retVal := this.PropPut(0x00000145, []interface{}{rhs})
	_= retVal
}

func (this *Options) MatchFuzzySpace() bool {
	retVal := this.PropGet(0x00000146, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetMatchFuzzySpace(rhs bool)  {
	retVal := this.PropPut(0x00000146, []interface{}{rhs})
	_= retVal
}

func (this *Options) ApplyFarEastFontsToAscii() bool {
	retVal := this.PropGet(0x00000147, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetApplyFarEastFontsToAscii(rhs bool)  {
	retVal := this.PropPut(0x00000147, []interface{}{rhs})
	_= retVal
}

func (this *Options) ConvertHighAnsiToFarEast() bool {
	retVal := this.PropGet(0x00000148, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetConvertHighAnsiToFarEast(rhs bool)  {
	retVal := this.PropPut(0x00000148, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintOddPagesInAscendingOrder() bool {
	retVal := this.PropGet(0x0000014a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintOddPagesInAscendingOrder(rhs bool)  {
	retVal := this.PropPut(0x0000014a, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintEvenPagesInAscendingOrder() bool {
	retVal := this.PropGet(0x0000014b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintEvenPagesInAscendingOrder(rhs bool)  {
	retVal := this.PropPut(0x0000014b, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultBorderColorIndex() int32 {
	retVal := this.PropGet(0x00000151, nil)
	return retVal.LValVal()
}

func (this *Options) SetDefaultBorderColorIndex(rhs int32)  {
	retVal := this.PropPut(0x00000151, []interface{}{rhs})
	_= retVal
}

func (this *Options) EnableMisusedWordsDictionary() bool {
	retVal := this.PropGet(0x00000152, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetEnableMisusedWordsDictionary(rhs bool)  {
	retVal := this.PropPut(0x00000152, []interface{}{rhs})
	_= retVal
}

func (this *Options) AllowCombinedAuxiliaryForms() bool {
	retVal := this.PropGet(0x00000153, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAllowCombinedAuxiliaryForms(rhs bool)  {
	retVal := this.PropPut(0x00000153, []interface{}{rhs})
	_= retVal
}

func (this *Options) HangulHanjaFastConversion() bool {
	retVal := this.PropGet(0x00000154, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetHangulHanjaFastConversion(rhs bool)  {
	retVal := this.PropPut(0x00000154, []interface{}{rhs})
	_= retVal
}

func (this *Options) CheckHangulEndings() bool {
	retVal := this.PropGet(0x00000155, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetCheckHangulEndings(rhs bool)  {
	retVal := this.PropPut(0x00000155, []interface{}{rhs})
	_= retVal
}

func (this *Options) EnableHangulHanjaRecentOrdering() bool {
	retVal := this.PropGet(0x00000156, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetEnableHangulHanjaRecentOrdering(rhs bool)  {
	retVal := this.PropPut(0x00000156, []interface{}{rhs})
	_= retVal
}

func (this *Options) MultipleWordConversionsMode() int32 {
	retVal := this.PropGet(0x00000157, nil)
	return retVal.LValVal()
}

func (this *Options) SetMultipleWordConversionsMode(rhs int32)  {
	retVal := this.PropPut(0x00000157, []interface{}{rhs})
	_= retVal
}

var Options_SetWPHelpOptions_OptArgs= []string{
	"CommandKeyHelp", "DocNavigationKeys", "MouseSimulation", "DemoGuidance", 
	"DemoSpeed", "HelpType", 
}

func (this *Options) SetWPHelpOptions(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Options_SetWPHelpOptions_OptArgs, optArgs)
	retVal := this.Call(0x0000014d, nil, optArgs...)
	_= retVal
}

func (this *Options) DefaultBorderColor() int32 {
	retVal := this.PropGet(0x00000158, nil)
	return retVal.LValVal()
}

func (this *Options) SetDefaultBorderColor(rhs int32)  {
	retVal := this.PropPut(0x00000158, []interface{}{rhs})
	_= retVal
}

func (this *Options) AllowPixelUnits() bool {
	retVal := this.PropGet(0x00000159, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAllowPixelUnits(rhs bool)  {
	retVal := this.PropPut(0x00000159, []interface{}{rhs})
	_= retVal
}

func (this *Options) UseCharacterUnit() bool {
	retVal := this.PropGet(0x0000015a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetUseCharacterUnit(rhs bool)  {
	retVal := this.PropPut(0x0000015a, []interface{}{rhs})
	_= retVal
}

func (this *Options) AllowCompoundNounProcessing() bool {
	retVal := this.PropGet(0x0000015b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAllowCompoundNounProcessing(rhs bool)  {
	retVal := this.PropPut(0x0000015b, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoKeyboardSwitching() bool {
	retVal := this.PropGet(0x0000018f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoKeyboardSwitching(rhs bool)  {
	retVal := this.PropPut(0x0000018f, []interface{}{rhs})
	_= retVal
}

func (this *Options) DocumentViewDirection() int32 {
	retVal := this.PropGet(0x00000190, nil)
	return retVal.LValVal()
}

func (this *Options) SetDocumentViewDirection(rhs int32)  {
	retVal := this.PropPut(0x00000190, []interface{}{rhs})
	_= retVal
}

func (this *Options) ArabicNumeral() int32 {
	retVal := this.PropGet(0x00000191, nil)
	return retVal.LValVal()
}

func (this *Options) SetArabicNumeral(rhs int32)  {
	retVal := this.PropPut(0x00000191, []interface{}{rhs})
	_= retVal
}

func (this *Options) MonthNames() int32 {
	retVal := this.PropGet(0x00000192, nil)
	return retVal.LValVal()
}

func (this *Options) SetMonthNames(rhs int32)  {
	retVal := this.PropPut(0x00000192, []interface{}{rhs})
	_= retVal
}

func (this *Options) CursorMovement() int32 {
	retVal := this.PropGet(0x00000193, nil)
	return retVal.LValVal()
}

func (this *Options) SetCursorMovement(rhs int32)  {
	retVal := this.PropPut(0x00000193, []interface{}{rhs})
	_= retVal
}

func (this *Options) VisualSelection() int32 {
	retVal := this.PropGet(0x00000194, nil)
	return retVal.LValVal()
}

func (this *Options) SetVisualSelection(rhs int32)  {
	retVal := this.PropPut(0x00000194, []interface{}{rhs})
	_= retVal
}

func (this *Options) ShowDiacritics() bool {
	retVal := this.PropGet(0x00000195, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetShowDiacritics(rhs bool)  {
	retVal := this.PropPut(0x00000195, []interface{}{rhs})
	_= retVal
}

func (this *Options) ShowControlCharacters() bool {
	retVal := this.PropGet(0x00000196, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetShowControlCharacters(rhs bool)  {
	retVal := this.PropPut(0x00000196, []interface{}{rhs})
	_= retVal
}

func (this *Options) AddControlCharacters() bool {
	retVal := this.PropGet(0x00000197, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAddControlCharacters(rhs bool)  {
	retVal := this.PropPut(0x00000197, []interface{}{rhs})
	_= retVal
}

func (this *Options) AddBiDirectionalMarksWhenSavingTextFile() bool {
	retVal := this.PropGet(0x00000198, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAddBiDirectionalMarksWhenSavingTextFile(rhs bool)  {
	retVal := this.PropPut(0x00000198, []interface{}{rhs})
	_= retVal
}

func (this *Options) StrictInitialAlefHamza() bool {
	retVal := this.PropGet(0x00000199, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetStrictInitialAlefHamza(rhs bool)  {
	retVal := this.PropPut(0x00000199, []interface{}{rhs})
	_= retVal
}

func (this *Options) StrictFinalYaa() bool {
	retVal := this.PropGet(0x0000019a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetStrictFinalYaa(rhs bool)  {
	retVal := this.PropPut(0x0000019a, []interface{}{rhs})
	_= retVal
}

func (this *Options) HebrewMode() int32 {
	retVal := this.PropGet(0x0000019b, nil)
	return retVal.LValVal()
}

func (this *Options) SetHebrewMode(rhs int32)  {
	retVal := this.PropPut(0x0000019b, []interface{}{rhs})
	_= retVal
}

func (this *Options) ArabicMode() int32 {
	retVal := this.PropGet(0x0000019c, nil)
	return retVal.LValVal()
}

func (this *Options) SetArabicMode(rhs int32)  {
	retVal := this.PropPut(0x0000019c, []interface{}{rhs})
	_= retVal
}

func (this *Options) AllowClickAndTypeMouse() bool {
	retVal := this.PropGet(0x0000019d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAllowClickAndTypeMouse(rhs bool)  {
	retVal := this.PropPut(0x0000019d, []interface{}{rhs})
	_= retVal
}

func (this *Options) UseGermanSpellingReform() bool {
	retVal := this.PropGet(0x0000019f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetUseGermanSpellingReform(rhs bool)  {
	retVal := this.PropPut(0x0000019f, []interface{}{rhs})
	_= retVal
}

func (this *Options) InterpretHighAnsi() int32 {
	retVal := this.PropGet(0x000001a2, nil)
	return retVal.LValVal()
}

func (this *Options) SetInterpretHighAnsi(rhs int32)  {
	retVal := this.PropPut(0x000001a2, []interface{}{rhs})
	_= retVal
}

func (this *Options) AddHebDoubleQuote() bool {
	retVal := this.PropGet(0x000001a3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAddHebDoubleQuote(rhs bool)  {
	retVal := this.PropPut(0x000001a3, []interface{}{rhs})
	_= retVal
}

func (this *Options) UseDiffDiacColor() bool {
	retVal := this.PropGet(0x000001a4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetUseDiffDiacColor(rhs bool)  {
	retVal := this.PropPut(0x000001a4, []interface{}{rhs})
	_= retVal
}

func (this *Options) DiacriticColorVal() int32 {
	retVal := this.PropGet(0x000001a5, nil)
	return retVal.LValVal()
}

func (this *Options) SetDiacriticColorVal(rhs int32)  {
	retVal := this.PropPut(0x000001a5, []interface{}{rhs})
	_= retVal
}

func (this *Options) OptimizeForWord97byDefault() bool {
	retVal := this.PropGet(0x000001a7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetOptimizeForWord97byDefault(rhs bool)  {
	retVal := this.PropPut(0x000001a7, []interface{}{rhs})
	_= retVal
}

func (this *Options) LocalNetworkFile() bool {
	retVal := this.PropGet(0x000001a8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetLocalNetworkFile(rhs bool)  {
	retVal := this.PropPut(0x000001a8, []interface{}{rhs})
	_= retVal
}

func (this *Options) TypeNReplace() bool {
	retVal := this.PropGet(0x000001a9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetTypeNReplace(rhs bool)  {
	retVal := this.PropPut(0x000001a9, []interface{}{rhs})
	_= retVal
}

func (this *Options) SequenceCheck() bool {
	retVal := this.PropGet(0x000001aa, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSequenceCheck(rhs bool)  {
	retVal := this.PropPut(0x000001aa, []interface{}{rhs})
	_= retVal
}

func (this *Options) BackgroundOpen() bool {
	retVal := this.PropGet(0x000001ab, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetBackgroundOpen(rhs bool)  {
	retVal := this.PropPut(0x000001ab, []interface{}{rhs})
	_= retVal
}

func (this *Options) DisableFeaturesbyDefault() bool {
	retVal := this.PropGet(0x000001ac, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetDisableFeaturesbyDefault(rhs bool)  {
	retVal := this.PropPut(0x000001ac, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteAdjustWordSpacing() bool {
	retVal := this.PropGet(0x000001ad, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPasteAdjustWordSpacing(rhs bool)  {
	retVal := this.PropPut(0x000001ad, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteAdjustParagraphSpacing() bool {
	retVal := this.PropGet(0x000001ae, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPasteAdjustParagraphSpacing(rhs bool)  {
	retVal := this.PropPut(0x000001ae, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteAdjustTableFormatting() bool {
	retVal := this.PropGet(0x000001af, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPasteAdjustTableFormatting(rhs bool)  {
	retVal := this.PropPut(0x000001af, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteSmartStyleBehavior() bool {
	retVal := this.PropGet(0x000001b0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPasteSmartStyleBehavior(rhs bool)  {
	retVal := this.PropPut(0x000001b0, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteMergeFromPPT() bool {
	retVal := this.PropGet(0x000001b1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPasteMergeFromPPT(rhs bool)  {
	retVal := this.PropPut(0x000001b1, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteMergeFromXL() bool {
	retVal := this.PropGet(0x000001b2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPasteMergeFromXL(rhs bool)  {
	retVal := this.PropPut(0x000001b2, []interface{}{rhs})
	_= retVal
}

func (this *Options) CtrlClickHyperlinkToOpen() bool {
	retVal := this.PropGet(0x000001b3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetCtrlClickHyperlinkToOpen(rhs bool)  {
	retVal := this.PropPut(0x000001b3, []interface{}{rhs})
	_= retVal
}

func (this *Options) PictureWrapType() int32 {
	retVal := this.PropGet(0x000001b4, nil)
	return retVal.LValVal()
}

func (this *Options) SetPictureWrapType(rhs int32)  {
	retVal := this.PropPut(0x000001b4, []interface{}{rhs})
	_= retVal
}

func (this *Options) DisableFeaturesIntroducedAfterbyDefault() int32 {
	retVal := this.PropGet(0x000001b5, nil)
	return retVal.LValVal()
}

func (this *Options) SetDisableFeaturesIntroducedAfterbyDefault(rhs int32)  {
	retVal := this.PropPut(0x000001b5, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteSmartCutPaste() bool {
	retVal := this.PropGet(0x000001b6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPasteSmartCutPaste(rhs bool)  {
	retVal := this.PropPut(0x000001b6, []interface{}{rhs})
	_= retVal
}

func (this *Options) DisplayPasteOptions() bool {
	retVal := this.PropGet(0x000001b7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetDisplayPasteOptions(rhs bool)  {
	retVal := this.PropPut(0x000001b7, []interface{}{rhs})
	_= retVal
}

func (this *Options) PromptUpdateStyle() bool {
	retVal := this.PropGet(0x000001b9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPromptUpdateStyle(rhs bool)  {
	retVal := this.PropPut(0x000001b9, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultEPostageApp() string {
	retVal := this.PropGet(0x000001ba, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Options) SetDefaultEPostageApp(rhs string)  {
	retVal := this.PropPut(0x000001ba, []interface{}{rhs})
	_= retVal
}

func (this *Options) DefaultTextEncoding() int32 {
	retVal := this.PropGet(0x000001bb, nil)
	return retVal.LValVal()
}

func (this *Options) SetDefaultTextEncoding(rhs int32)  {
	retVal := this.PropPut(0x000001bb, []interface{}{rhs})
	_= retVal
}

func (this *Options) LabelSmartTags() bool {
	retVal := this.PropGet(0x000001bc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetLabelSmartTags(rhs bool)  {
	retVal := this.PropPut(0x000001bc, []interface{}{rhs})
	_= retVal
}

func (this *Options) DisplaySmartTagButtons() bool {
	retVal := this.PropGet(0x000001bd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetDisplaySmartTagButtons(rhs bool)  {
	retVal := this.PropPut(0x000001bd, []interface{}{rhs})
	_= retVal
}

func (this *Options) WarnBeforeSavingPrintingSendingMarkup() bool {
	retVal := this.PropGet(0x000001be, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetWarnBeforeSavingPrintingSendingMarkup(rhs bool)  {
	retVal := this.PropPut(0x000001be, []interface{}{rhs})
	_= retVal
}

func (this *Options) StoreRSIDOnSave() bool {
	retVal := this.PropGet(0x000001bf, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetStoreRSIDOnSave(rhs bool)  {
	retVal := this.PropPut(0x000001bf, []interface{}{rhs})
	_= retVal
}

func (this *Options) ShowFormatError() bool {
	retVal := this.PropGet(0x000001c0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetShowFormatError(rhs bool)  {
	retVal := this.PropPut(0x000001c0, []interface{}{rhs})
	_= retVal
}

func (this *Options) FormatScanning() bool {
	retVal := this.PropGet(0x000001c1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetFormatScanning(rhs bool)  {
	retVal := this.PropPut(0x000001c1, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteMergeLists() bool {
	retVal := this.PropGet(0x000001c2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPasteMergeLists(rhs bool)  {
	retVal := this.PropPut(0x000001c2, []interface{}{rhs})
	_= retVal
}

func (this *Options) AutoCreateNewDrawings() bool {
	retVal := this.PropGet(0x000001c3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAutoCreateNewDrawings(rhs bool)  {
	retVal := this.PropPut(0x000001c3, []interface{}{rhs})
	_= retVal
}

func (this *Options) SmartParaSelection() bool {
	retVal := this.PropGet(0x000001c4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSmartParaSelection(rhs bool)  {
	retVal := this.PropPut(0x000001c4, []interface{}{rhs})
	_= retVal
}

func (this *Options) RevisionsBalloonPrintOrientation() int32 {
	retVal := this.PropGet(0x000001c5, nil)
	return retVal.LValVal()
}

func (this *Options) SetRevisionsBalloonPrintOrientation(rhs int32)  {
	retVal := this.PropPut(0x000001c5, []interface{}{rhs})
	_= retVal
}

func (this *Options) CommentsColor() int32 {
	retVal := this.PropGet(0x000001c6, nil)
	return retVal.LValVal()
}

func (this *Options) SetCommentsColor(rhs int32)  {
	retVal := this.PropPut(0x000001c6, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintXMLTag() bool {
	retVal := this.PropGet(0x000001c7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintXMLTag(rhs bool)  {
	retVal := this.PropPut(0x000001c7, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrintBackgrounds() bool {
	retVal := this.PropGet(0x000001c8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrintBackgrounds(rhs bool)  {
	retVal := this.PropPut(0x000001c8, []interface{}{rhs})
	_= retVal
}

func (this *Options) AllowReadingMode() bool {
	retVal := this.PropGet(0x000001c9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAllowReadingMode(rhs bool)  {
	retVal := this.PropPut(0x000001c9, []interface{}{rhs})
	_= retVal
}

func (this *Options) ShowMarkupOpenSave() bool {
	retVal := this.PropGet(0x000001ca, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetShowMarkupOpenSave(rhs bool)  {
	retVal := this.PropPut(0x000001ca, []interface{}{rhs})
	_= retVal
}

func (this *Options) SmartCursoring() bool {
	retVal := this.PropGet(0x000001cb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetSmartCursoring(rhs bool)  {
	retVal := this.PropPut(0x000001cb, []interface{}{rhs})
	_= retVal
}

func (this *Options) MoveToTextMark() int32 {
	retVal := this.PropGet(0x000001cc, nil)
	return retVal.LValVal()
}

func (this *Options) SetMoveToTextMark(rhs int32)  {
	retVal := this.PropPut(0x000001cc, []interface{}{rhs})
	_= retVal
}

func (this *Options) MoveFromTextMark() int32 {
	retVal := this.PropGet(0x000001cd, nil)
	return retVal.LValVal()
}

func (this *Options) SetMoveFromTextMark(rhs int32)  {
	retVal := this.PropPut(0x000001cd, []interface{}{rhs})
	_= retVal
}

func (this *Options) BibliographyStyle() string {
	retVal := this.PropGet(0x000001ce, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Options) SetBibliographyStyle(rhs string)  {
	retVal := this.PropPut(0x000001ce, []interface{}{rhs})
	_= retVal
}

func (this *Options) BibliographySort() string {
	retVal := this.PropGet(0x000001cf, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Options) SetBibliographySort(rhs string)  {
	retVal := this.PropPut(0x000001cf, []interface{}{rhs})
	_= retVal
}

func (this *Options) InsertedCellColor() int32 {
	retVal := this.PropGet(0x000001d0, nil)
	return retVal.LValVal()
}

func (this *Options) SetInsertedCellColor(rhs int32)  {
	retVal := this.PropPut(0x000001d0, []interface{}{rhs})
	_= retVal
}

func (this *Options) DeletedCellColor() int32 {
	retVal := this.PropGet(0x000001d1, nil)
	return retVal.LValVal()
}

func (this *Options) SetDeletedCellColor(rhs int32)  {
	retVal := this.PropPut(0x000001d1, []interface{}{rhs})
	_= retVal
}

func (this *Options) MergedCellColor() int32 {
	retVal := this.PropGet(0x000001d2, nil)
	return retVal.LValVal()
}

func (this *Options) SetMergedCellColor(rhs int32)  {
	retVal := this.PropPut(0x000001d2, []interface{}{rhs})
	_= retVal
}

func (this *Options) SplitCellColor() int32 {
	retVal := this.PropGet(0x000001d3, nil)
	return retVal.LValVal()
}

func (this *Options) SetSplitCellColor(rhs int32)  {
	retVal := this.PropPut(0x000001d3, []interface{}{rhs})
	_= retVal
}

func (this *Options) ShowSelectionFloaties() bool {
	retVal := this.PropGet(0x000001d4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetShowSelectionFloaties(rhs bool)  {
	retVal := this.PropPut(0x000001d4, []interface{}{rhs})
	_= retVal
}

func (this *Options) ShowMenuFloaties() bool {
	retVal := this.PropGet(0x000001d5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetShowMenuFloaties(rhs bool)  {
	retVal := this.PropPut(0x000001d5, []interface{}{rhs})
	_= retVal
}

func (this *Options) ShowDevTools() bool {
	retVal := this.PropGet(0x000001d6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetShowDevTools(rhs bool)  {
	retVal := this.PropPut(0x000001d6, []interface{}{rhs})
	_= retVal
}

func (this *Options) EnableLivePreview() bool {
	retVal := this.PropGet(0x000001d8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetEnableLivePreview(rhs bool)  {
	retVal := this.PropPut(0x000001d8, []interface{}{rhs})
	_= retVal
}

func (this *Options) OMathAutoBuildUp() bool {
	retVal := this.PropGet(0x000001da, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetOMathAutoBuildUp(rhs bool)  {
	retVal := this.PropPut(0x000001da, []interface{}{rhs})
	_= retVal
}

func (this *Options) AlwaysUseClearType() bool {
	retVal := this.PropGet(0x000001dc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAlwaysUseClearType(rhs bool)  {
	retVal := this.PropPut(0x000001dc, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteFormatWithinDocument() int32 {
	retVal := this.PropGet(0x000001dd, nil)
	return retVal.LValVal()
}

func (this *Options) SetPasteFormatWithinDocument(rhs int32)  {
	retVal := this.PropPut(0x000001dd, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteFormatBetweenDocuments() int32 {
	retVal := this.PropGet(0x000001de, nil)
	return retVal.LValVal()
}

func (this *Options) SetPasteFormatBetweenDocuments(rhs int32)  {
	retVal := this.PropPut(0x000001de, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteFormatBetweenStyledDocuments() int32 {
	retVal := this.PropGet(0x000001df, nil)
	return retVal.LValVal()
}

func (this *Options) SetPasteFormatBetweenStyledDocuments(rhs int32)  {
	retVal := this.PropPut(0x000001df, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteFormatFromExternalSource() int32 {
	retVal := this.PropGet(0x000001e0, nil)
	return retVal.LValVal()
}

func (this *Options) SetPasteFormatFromExternalSource(rhs int32)  {
	retVal := this.PropPut(0x000001e0, []interface{}{rhs})
	_= retVal
}

func (this *Options) PasteOptionKeepBulletsAndNumbers() bool {
	retVal := this.PropGet(0x000001e1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPasteOptionKeepBulletsAndNumbers(rhs bool)  {
	retVal := this.PropPut(0x000001e1, []interface{}{rhs})
	_= retVal
}

func (this *Options) INSKeyForOvertype() bool {
	retVal := this.PropGet(0x000001e2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetINSKeyForOvertype(rhs bool)  {
	retVal := this.PropPut(0x000001e2, []interface{}{rhs})
	_= retVal
}

func (this *Options) RepeatWord() bool {
	retVal := this.PropGet(0x000001e3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetRepeatWord(rhs bool)  {
	retVal := this.PropPut(0x000001e3, []interface{}{rhs})
	_= retVal
}

func (this *Options) FrenchReform() int32 {
	retVal := this.PropGet(0x000001e4, nil)
	return retVal.LValVal()
}

func (this *Options) SetFrenchReform(rhs int32)  {
	retVal := this.PropPut(0x000001e4, []interface{}{rhs})
	_= retVal
}

func (this *Options) ContextualSpeller() bool {
	retVal := this.PropGet(0x000001e5, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetContextualSpeller(rhs bool)  {
	retVal := this.PropPut(0x000001e5, []interface{}{rhs})
	_= retVal
}

func (this *Options) MoveToTextColor() int32 {
	retVal := this.PropGet(0x000001e6, nil)
	return retVal.LValVal()
}

func (this *Options) SetMoveToTextColor(rhs int32)  {
	retVal := this.PropPut(0x000001e6, []interface{}{rhs})
	_= retVal
}

func (this *Options) MoveFromTextColor() int32 {
	retVal := this.PropGet(0x000001e7, nil)
	return retVal.LValVal()
}

func (this *Options) SetMoveFromTextColor(rhs int32)  {
	retVal := this.PropPut(0x000001e7, []interface{}{rhs})
	_= retVal
}

func (this *Options) OMathCopyLF() bool {
	retVal := this.PropGet(0x000001e8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetOMathCopyLF(rhs bool)  {
	retVal := this.PropPut(0x000001e8, []interface{}{rhs})
	_= retVal
}

func (this *Options) UseNormalStyleForList() bool {
	retVal := this.PropGet(0x000001e9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetUseNormalStyleForList(rhs bool)  {
	retVal := this.PropPut(0x000001e9, []interface{}{rhs})
	_= retVal
}

func (this *Options) AllowOpenInDraftView() bool {
	retVal := this.PropGet(0x000001ea, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetAllowOpenInDraftView(rhs bool)  {
	retVal := this.PropPut(0x000001ea, []interface{}{rhs})
	_= retVal
}

func (this *Options) EnableLegacyIMEMode() bool {
	retVal := this.PropGet(0x000001ec, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetEnableLegacyIMEMode(rhs bool)  {
	retVal := this.PropPut(0x000001ec, []interface{}{rhs})
	_= retVal
}

func (this *Options) DoNotPromptForConvert() bool {
	retVal := this.PropGet(0x000001ed, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetDoNotPromptForConvert(rhs bool)  {
	retVal := this.PropPut(0x000001ed, []interface{}{rhs})
	_= retVal
}

func (this *Options) PrecisePositioning() bool {
	retVal := this.PropGet(0x000001ee, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetPrecisePositioning(rhs bool)  {
	retVal := this.PropPut(0x000001ee, []interface{}{rhs})
	_= retVal
}

func (this *Options) UpdateStyleListBehavior() int32 {
	retVal := this.PropGet(0x000001ef, nil)
	return retVal.LValVal()
}

func (this *Options) SetUpdateStyleListBehavior(rhs int32)  {
	retVal := this.PropPut(0x000001ef, []interface{}{rhs})
	_= retVal
}

func (this *Options) StrictTaaMarboota() bool {
	retVal := this.PropGet(0x000001f0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetStrictTaaMarboota(rhs bool)  {
	retVal := this.PropPut(0x000001f0, []interface{}{rhs})
	_= retVal
}

func (this *Options) StrictRussianE() bool {
	retVal := this.PropGet(0x000001f1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetStrictRussianE(rhs bool)  {
	retVal := this.PropPut(0x000001f1, []interface{}{rhs})
	_= retVal
}

func (this *Options) SpanishMode() int32 {
	retVal := this.PropGet(0x000001f2, nil)
	return retVal.LValVal()
}

func (this *Options) SetSpanishMode(rhs int32)  {
	retVal := this.PropPut(0x000001f2, []interface{}{rhs})
	_= retVal
}

func (this *Options) PortugalReform() int32 {
	retVal := this.PropGet(0x000001f5, nil)
	return retVal.LValVal()
}

func (this *Options) SetPortugalReform(rhs int32)  {
	retVal := this.PropPut(0x000001f5, []interface{}{rhs})
	_= retVal
}

func (this *Options) BrazilReform() int32 {
	retVal := this.PropGet(0x000001f6, nil)
	return retVal.LValVal()
}

func (this *Options) SetBrazilReform(rhs int32)  {
	retVal := this.PropPut(0x000001f6, []interface{}{rhs})
	_= retVal
}

func (this *Options) UpdateFieldsWithTrackedChangesAtPrint() bool {
	retVal := this.PropGet(0x000001f7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Options) SetUpdateFieldsWithTrackedChangesAtPrint(rhs bool)  {
	retVal := this.PropPut(0x000001f7, []interface{}{rhs})
	_= retVal
}

