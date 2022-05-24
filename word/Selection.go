package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020975-0000-0000-C000-000000000046
var IID_Selection = syscall.GUID{0x00020975, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Selection struct {
	ole.OleClient
}

func NewSelection(pDisp *win32.IDispatch, addRef bool, scoped bool) *Selection {
	 if pDisp == nil {
		return nil;
	}
	p := &Selection{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SelectionFromVar(v ole.Variant) *Selection {
	return NewSelection(v.IDispatch(), false, false)
}

func (this *Selection) IID() *syscall.GUID {
	return &IID_Selection
}

func (this *Selection) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Selection) Text() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Selection) SetText(rhs string)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *Selection) FormattedText() *Range {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Selection) SetFormattedText(rhs *Range)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *Selection) Start() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Selection) SetStart(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Selection) End() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Selection) SetEnd(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *Selection) Font() *Font {
	retVal, _ := this.PropGet(0x00000005, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *Selection) SetFont(rhs *Font)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *Selection) Type() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Selection) StoryType() int32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *Selection) Style() ole.Variant {
	retVal, _ := this.PropGet(0x00000008, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Selection) SetStyle(rhs *ole.Variant)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *Selection) Tables() *Tables {
	retVal, _ := this.PropGet(0x00000032, nil)
	return NewTables(retVal.IDispatch(), false, true)
}

func (this *Selection) Words() *Words {
	retVal, _ := this.PropGet(0x00000033, nil)
	return NewWords(retVal.IDispatch(), false, true)
}

func (this *Selection) Sentences() *Sentences {
	retVal, _ := this.PropGet(0x00000034, nil)
	return NewSentences(retVal.IDispatch(), false, true)
}

func (this *Selection) Characters() *Characters {
	retVal, _ := this.PropGet(0x00000035, nil)
	return NewCharacters(retVal.IDispatch(), false, true)
}

func (this *Selection) Footnotes() *Footnotes {
	retVal, _ := this.PropGet(0x00000036, nil)
	return NewFootnotes(retVal.IDispatch(), false, true)
}

func (this *Selection) Endnotes() *Endnotes {
	retVal, _ := this.PropGet(0x00000037, nil)
	return NewEndnotes(retVal.IDispatch(), false, true)
}

func (this *Selection) Comments() *Comments {
	retVal, _ := this.PropGet(0x00000038, nil)
	return NewComments(retVal.IDispatch(), false, true)
}

func (this *Selection) Cells() *Cells {
	retVal, _ := this.PropGet(0x00000039, nil)
	return NewCells(retVal.IDispatch(), false, true)
}

func (this *Selection) Sections() *Sections {
	retVal, _ := this.PropGet(0x0000003a, nil)
	return NewSections(retVal.IDispatch(), false, true)
}

func (this *Selection) Paragraphs() *Paragraphs {
	retVal, _ := this.PropGet(0x0000003b, nil)
	return NewParagraphs(retVal.IDispatch(), false, true)
}

func (this *Selection) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Selection) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Selection) Shading() *Shading {
	retVal, _ := this.PropGet(0x0000003d, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Selection) Fields() *Fields {
	retVal, _ := this.PropGet(0x00000040, nil)
	return NewFields(retVal.IDispatch(), false, true)
}

func (this *Selection) FormFields() *FormFields {
	retVal, _ := this.PropGet(0x00000041, nil)
	return NewFormFields(retVal.IDispatch(), false, true)
}

func (this *Selection) Frames() *Frames {
	retVal, _ := this.PropGet(0x00000042, nil)
	return NewFrames(retVal.IDispatch(), false, true)
}

func (this *Selection) ParagraphFormat() *ParagraphFormat {
	retVal, _ := this.PropGet(0x0000044e, nil)
	return NewParagraphFormat(retVal.IDispatch(), false, true)
}

func (this *Selection) SetParagraphFormat(rhs *ParagraphFormat)  {
	_ = this.PropPut(0x0000044e, []interface{}{rhs})
}

func (this *Selection) PageSetup() *PageSetup {
	retVal, _ := this.PropGet(0x0000044d, nil)
	return NewPageSetup(retVal.IDispatch(), false, true)
}

func (this *Selection) SetPageSetup(rhs *PageSetup)  {
	_ = this.PropPut(0x0000044d, []interface{}{rhs})
}

func (this *Selection) Bookmarks() *Bookmarks {
	retVal, _ := this.PropGet(0x0000004b, nil)
	return NewBookmarks(retVal.IDispatch(), false, true)
}

func (this *Selection) StoryLength() int32 {
	retVal, _ := this.PropGet(0x00000098, nil)
	return retVal.LValVal()
}

func (this *Selection) LanguageID() int32 {
	retVal, _ := this.PropGet(0x00000099, nil)
	return retVal.LValVal()
}

func (this *Selection) SetLanguageID(rhs int32)  {
	_ = this.PropPut(0x00000099, []interface{}{rhs})
}

func (this *Selection) LanguageIDFarEast() int32 {
	retVal, _ := this.PropGet(0x0000009a, nil)
	return retVal.LValVal()
}

func (this *Selection) SetLanguageIDFarEast(rhs int32)  {
	_ = this.PropPut(0x0000009a, []interface{}{rhs})
}

func (this *Selection) LanguageIDOther() int32 {
	retVal, _ := this.PropGet(0x0000009b, nil)
	return retVal.LValVal()
}

func (this *Selection) SetLanguageIDOther(rhs int32)  {
	_ = this.PropPut(0x0000009b, []interface{}{rhs})
}

func (this *Selection) Hyperlinks() *Hyperlinks {
	retVal, _ := this.PropGet(0x0000009c, nil)
	return NewHyperlinks(retVal.IDispatch(), false, true)
}

func (this *Selection) Columns() *Columns {
	retVal, _ := this.PropGet(0x0000012e, nil)
	return NewColumns(retVal.IDispatch(), false, true)
}

func (this *Selection) Rows() *Rows {
	retVal, _ := this.PropGet(0x0000012f, nil)
	return NewRows(retVal.IDispatch(), false, true)
}

func (this *Selection) HeaderFooter() *HeaderFooter {
	retVal, _ := this.PropGet(0x00000132, nil)
	return NewHeaderFooter(retVal.IDispatch(), false, true)
}

func (this *Selection) IsEndOfRowMark() bool {
	retVal, _ := this.PropGet(0x00000133, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) BookmarkID() int32 {
	retVal, _ := this.PropGet(0x00000134, nil)
	return retVal.LValVal()
}

func (this *Selection) PreviousBookmarkID() int32 {
	retVal, _ := this.PropGet(0x00000135, nil)
	return retVal.LValVal()
}

func (this *Selection) Find() *Find {
	retVal, _ := this.PropGet(0x00000106, nil)
	return NewFind(retVal.IDispatch(), false, true)
}

func (this *Selection) Range() *Range {
	retVal, _ := this.PropGet(0x00000190, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Selection) Information(type_ int32) ole.Variant {
	retVal, _ := this.PropGet(0x00000191, []interface{}{type_})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Selection) Flags() int32 {
	retVal, _ := this.PropGet(0x00000192, nil)
	return retVal.LValVal()
}

func (this *Selection) SetFlags(rhs int32)  {
	_ = this.PropPut(0x00000192, []interface{}{rhs})
}

func (this *Selection) Active() bool {
	retVal, _ := this.PropGet(0x00000193, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) StartIsActive() bool {
	retVal, _ := this.PropGet(0x00000194, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) SetStartIsActive(rhs bool)  {
	_ = this.PropPut(0x00000194, []interface{}{rhs})
}

func (this *Selection) IPAtEndOfLine() bool {
	retVal, _ := this.PropGet(0x00000195, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) ExtendMode() bool {
	retVal, _ := this.PropGet(0x00000196, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) SetExtendMode(rhs bool)  {
	_ = this.PropPut(0x00000196, []interface{}{rhs})
}

func (this *Selection) ColumnSelectMode() bool {
	retVal, _ := this.PropGet(0x00000197, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) SetColumnSelectMode(rhs bool)  {
	_ = this.PropPut(0x00000197, []interface{}{rhs})
}

func (this *Selection) Orientation() int32 {
	retVal, _ := this.PropGet(0x0000019a, nil)
	return retVal.LValVal()
}

func (this *Selection) SetOrientation(rhs int32)  {
	_ = this.PropPut(0x0000019a, []interface{}{rhs})
}

func (this *Selection) InlineShapes() *InlineShapes {
	retVal, _ := this.PropGet(0x0000019b, nil)
	return NewInlineShapes(retVal.IDispatch(), false, true)
}

func (this *Selection) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Selection) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Selection) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Selection) Document() *Document {
	retVal, _ := this.PropGet(0x000003eb, nil)
	return NewDocument(retVal.IDispatch(), false, true)
}

func (this *Selection) ShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000003ec, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *Selection) Select()  {
	retVal, _ := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *Selection) SetRange(start int32, end int32)  {
	retVal, _ := this.Call(0x00000064, []interface{}{start, end})
	_= retVal
}

var Selection_Collapse_OptArgs= []string{
	"Direction", 
}

func (this *Selection) Collapse(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_Collapse_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, nil, optArgs...)
	_= retVal
}

func (this *Selection) InsertBefore(text string)  {
	retVal, _ := this.Call(0x00000066, []interface{}{text})
	_= retVal
}

func (this *Selection) InsertAfter(text string)  {
	retVal, _ := this.Call(0x00000068, []interface{}{text})
	_= retVal
}

var Selection_Next_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Selection) Next(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Selection_Next_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000069, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

var Selection_Previous_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Selection) Previous(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Selection_Previous_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006a, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

var Selection_StartOf_OptArgs= []string{
	"Unit", "Extend", 
}

func (this *Selection) StartOf(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_StartOf_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006b, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_EndOf_OptArgs= []string{
	"Unit", "Extend", 
}

func (this *Selection) EndOf(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_EndOf_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006c, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_Move_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Selection) Move(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_Move_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006d, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveStart_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Selection) MoveStart(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveStart_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006e, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveEnd_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Selection) MoveEnd(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveEnd_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006f, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveWhile_OptArgs= []string{
	"Count", 
}

func (this *Selection) MoveWhile(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveWhile_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000070, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveStartWhile_OptArgs= []string{
	"Count", 
}

func (this *Selection) MoveStartWhile(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveStartWhile_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000071, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveEndWhile_OptArgs= []string{
	"Count", 
}

func (this *Selection) MoveEndWhile(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveEndWhile_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000072, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveUntil_OptArgs= []string{
	"Count", 
}

func (this *Selection) MoveUntil(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveUntil_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000073, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveStartUntil_OptArgs= []string{
	"Count", 
}

func (this *Selection) MoveStartUntil(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveStartUntil_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000074, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveEndUntil_OptArgs= []string{
	"Count", 
}

func (this *Selection) MoveEndUntil(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveEndUntil_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000075, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

func (this *Selection) Cut()  {
	retVal, _ := this.Call(0x00000077, nil)
	_= retVal
}

func (this *Selection) Copy()  {
	retVal, _ := this.Call(0x00000078, nil)
	_= retVal
}

func (this *Selection) Paste()  {
	retVal, _ := this.Call(0x00000079, nil)
	_= retVal
}

var Selection_InsertBreak_OptArgs= []string{
	"Type", 
}

func (this *Selection) InsertBreak(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertBreak_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000007a, nil, optArgs...)
	_= retVal
}

var Selection_InsertFile_OptArgs= []string{
	"Range", "ConfirmConversions", "Link", "Attachment", 
}

func (this *Selection) InsertFile(fileName string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertFile_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000007b, []interface{}{fileName}, optArgs...)
	_= retVal
}

func (this *Selection) InStory(range_ *Range) bool {
	retVal, _ := this.Call(0x0000007d, []interface{}{range_})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) InRange(range_ *Range) bool {
	retVal, _ := this.Call(0x0000007e, []interface{}{range_})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Selection_Delete_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Selection) Delete(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_Delete_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000007f, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_Expand_OptArgs= []string{
	"Unit", 
}

func (this *Selection) Expand(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_Expand_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000081, nil, optArgs...)
	return retVal.LValVal()
}

func (this *Selection) InsertParagraph()  {
	retVal, _ := this.Call(0x000000a0, nil)
	_= retVal
}

func (this *Selection) InsertParagraphAfter()  {
	retVal, _ := this.Call(0x000000a1, nil)
	_= retVal
}

var Selection_ConvertToTableOld_OptArgs= []string{
	"Separator", "NumRows", "NumColumns", "InitialColumnWidth", 
	"Format", "ApplyBorders", "ApplyShading", "ApplyFont", 
	"ApplyColor", "ApplyHeadingRows", "ApplyLastRow", "ApplyFirstColumn", 
	"ApplyLastColumn", "AutoFit", 
}

func (this *Selection) ConvertToTableOld(optArgs ...interface{}) *Table {
	optArgs = ole.ProcessOptArgs(Selection_ConvertToTableOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000a2, nil, optArgs...)
	return NewTable(retVal.IDispatch(), false, true)
}

var Selection_InsertDateTimeOld_OptArgs= []string{
	"DateTimeFormat", "InsertAsField", "InsertAsFullWidth", 
}

func (this *Selection) InsertDateTimeOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertDateTimeOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000a3, nil, optArgs...)
	_= retVal
}

var Selection_InsertSymbol_OptArgs= []string{
	"Font", "Unicode", "Bias", 
}

func (this *Selection) InsertSymbol(characterNumber int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertSymbol_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000a4, []interface{}{characterNumber}, optArgs...)
	_= retVal
}

var Selection_InsertCrossReference_2002_OptArgs= []string{
	"InsertAsHyperlink", "IncludePosition", 
}

func (this *Selection) InsertCrossReference_2002(referenceType *ole.Variant, referenceKind int32, referenceItem *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertCrossReference_2002_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000a5, []interface{}{referenceType, referenceKind, referenceItem}, optArgs...)
	_= retVal
}

var Selection_InsertCaptionXP_OptArgs= []string{
	"Title", "TitleAutoText", "Position", 
}

func (this *Selection) InsertCaptionXP(label *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertCaptionXP_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000a6, []interface{}{label}, optArgs...)
	_= retVal
}

func (this *Selection) CopyAsPicture()  {
	retVal, _ := this.Call(0x000000a7, nil)
	_= retVal
}

var Selection_SortOld_OptArgs= []string{
	"ExcludeHeader", "FieldNumber", "SortFieldType", "SortOrder", 
	"FieldNumber2", "SortFieldType2", "SortOrder2", "FieldNumber3", 
	"SortFieldType3", "SortOrder3", "SortColumn", "Separator", 
	"CaseSensitive", "LanguageID", 
}

func (this *Selection) SortOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_SortOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000a8, nil, optArgs...)
	_= retVal
}

func (this *Selection) SortAscending()  {
	retVal, _ := this.Call(0x000000a9, nil)
	_= retVal
}

func (this *Selection) SortDescending()  {
	retVal, _ := this.Call(0x000000aa, nil)
	_= retVal
}

func (this *Selection) IsEqual(range_ *Range) bool {
	retVal, _ := this.Call(0x000000ab, []interface{}{range_})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) Calculate() float32 {
	retVal, _ := this.Call(0x000000ac, nil)
	return retVal.FltValVal()
}

var Selection_GoTo_OptArgs= []string{
	"What", "Which", "Count", "Name", 
}

func (this *Selection) GoTo(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Selection_GoTo_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000ad, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Selection) GoToNext(what int32) *Range {
	retVal, _ := this.Call(0x000000ae, []interface{}{what})
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Selection) GoToPrevious(what int32) *Range {
	retVal, _ := this.Call(0x000000af, []interface{}{what})
	return NewRange(retVal.IDispatch(), false, true)
}

var Selection_PasteSpecial_OptArgs= []string{
	"IconIndex", "Link", "Placement", "DisplayAsIcon", 
	"DataType", "IconFileName", "IconLabel", 
}

func (this *Selection) PasteSpecial(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_PasteSpecial_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000b0, nil, optArgs...)
	_= retVal
}

func (this *Selection) PreviousField() *Field {
	retVal, _ := this.Call(0x000000b1, nil)
	return NewField(retVal.IDispatch(), false, true)
}

func (this *Selection) NextField() *Field {
	retVal, _ := this.Call(0x000000b2, nil)
	return NewField(retVal.IDispatch(), false, true)
}

func (this *Selection) InsertParagraphBefore()  {
	retVal, _ := this.Call(0x000000d4, nil)
	_= retVal
}

var Selection_InsertCells_OptArgs= []string{
	"ShiftCells", 
}

func (this *Selection) InsertCells(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertCells_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000d6, nil, optArgs...)
	_= retVal
}

var Selection_Extend_OptArgs= []string{
	"Character", 
}

func (this *Selection) Extend(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_Extend_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000012c, nil, optArgs...)
	_= retVal
}

func (this *Selection) Shrink()  {
	retVal, _ := this.Call(0x0000012d, nil)
	_= retVal
}

var Selection_MoveLeft_OptArgs= []string{
	"Unit", "Count", "Extend", 
}

func (this *Selection) MoveLeft(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveLeft_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f4, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveRight_OptArgs= []string{
	"Unit", "Count", "Extend", 
}

func (this *Selection) MoveRight(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveRight_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f5, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveUp_OptArgs= []string{
	"Unit", "Count", "Extend", 
}

func (this *Selection) MoveUp(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveUp_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f6, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_MoveDown_OptArgs= []string{
	"Unit", "Count", "Extend", 
}

func (this *Selection) MoveDown(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_MoveDown_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f7, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_HomeKey_OptArgs= []string{
	"Unit", "Extend", 
}

func (this *Selection) HomeKey(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_HomeKey_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f8, nil, optArgs...)
	return retVal.LValVal()
}

var Selection_EndKey_OptArgs= []string{
	"Unit", "Extend", 
}

func (this *Selection) EndKey(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Selection_EndKey_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f9, nil, optArgs...)
	return retVal.LValVal()
}

func (this *Selection) EscapeKey()  {
	retVal, _ := this.Call(0x000001fa, nil)
	_= retVal
}

func (this *Selection) TypeText(text string)  {
	retVal, _ := this.Call(0x000001fb, []interface{}{text})
	_= retVal
}

func (this *Selection) CopyFormat()  {
	retVal, _ := this.Call(0x000001fd, nil)
	_= retVal
}

func (this *Selection) PasteFormat()  {
	retVal, _ := this.Call(0x000001fe, nil)
	_= retVal
}

func (this *Selection) TypeParagraph()  {
	retVal, _ := this.Call(0x00000200, nil)
	_= retVal
}

func (this *Selection) TypeBackspace()  {
	retVal, _ := this.Call(0x00000201, nil)
	_= retVal
}

func (this *Selection) NextSubdocument()  {
	retVal, _ := this.Call(0x00000202, nil)
	_= retVal
}

func (this *Selection) PreviousSubdocument()  {
	retVal, _ := this.Call(0x00000203, nil)
	_= retVal
}

func (this *Selection) SelectColumn()  {
	retVal, _ := this.Call(0x00000204, nil)
	_= retVal
}

func (this *Selection) SelectCurrentFont()  {
	retVal, _ := this.Call(0x00000205, nil)
	_= retVal
}

func (this *Selection) SelectCurrentAlignment()  {
	retVal, _ := this.Call(0x00000206, nil)
	_= retVal
}

func (this *Selection) SelectCurrentSpacing()  {
	retVal, _ := this.Call(0x00000207, nil)
	_= retVal
}

func (this *Selection) SelectCurrentIndent()  {
	retVal, _ := this.Call(0x00000208, nil)
	_= retVal
}

func (this *Selection) SelectCurrentTabs()  {
	retVal, _ := this.Call(0x00000209, nil)
	_= retVal
}

func (this *Selection) SelectCurrentColor()  {
	retVal, _ := this.Call(0x0000020a, nil)
	_= retVal
}

func (this *Selection) CreateTextbox()  {
	retVal, _ := this.Call(0x0000020b, nil)
	_= retVal
}

func (this *Selection) WholeStory()  {
	retVal, _ := this.Call(0x0000020c, nil)
	_= retVal
}

func (this *Selection) SelectRow()  {
	retVal, _ := this.Call(0x0000020d, nil)
	_= retVal
}

func (this *Selection) SplitTable()  {
	retVal, _ := this.Call(0x0000020e, nil)
	_= retVal
}

var Selection_InsertRows_OptArgs= []string{
	"NumRows", 
}

func (this *Selection) InsertRows(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertRows_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000210, nil, optArgs...)
	_= retVal
}

func (this *Selection) InsertColumns()  {
	retVal, _ := this.Call(0x00000211, nil)
	_= retVal
}

var Selection_InsertFormula_OptArgs= []string{
	"Formula", "NumberFormat", 
}

func (this *Selection) InsertFormula(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertFormula_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000212, nil, optArgs...)
	_= retVal
}

var Selection_NextRevision_OptArgs= []string{
	"Wrap", 
}

func (this *Selection) NextRevision(optArgs ...interface{}) *Revision {
	optArgs = ole.ProcessOptArgs(Selection_NextRevision_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000213, nil, optArgs...)
	return NewRevision(retVal.IDispatch(), false, true)
}

var Selection_PreviousRevision_OptArgs= []string{
	"Wrap", 
}

func (this *Selection) PreviousRevision(optArgs ...interface{}) *Revision {
	optArgs = ole.ProcessOptArgs(Selection_PreviousRevision_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000214, nil, optArgs...)
	return NewRevision(retVal.IDispatch(), false, true)
}

func (this *Selection) PasteAsNestedTable()  {
	retVal, _ := this.Call(0x00000215, nil)
	_= retVal
}

func (this *Selection) CreateAutoTextEntry(name string, styleName string) *AutoTextEntry {
	retVal, _ := this.Call(0x00000216, []interface{}{name, styleName})
	return NewAutoTextEntry(retVal.IDispatch(), false, true)
}

func (this *Selection) DetectLanguage()  {
	retVal, _ := this.Call(0x00000217, nil)
	_= retVal
}

func (this *Selection) SelectCell()  {
	retVal, _ := this.Call(0x00000218, nil)
	_= retVal
}

var Selection_InsertRowsBelow_OptArgs= []string{
	"NumRows", 
}

func (this *Selection) InsertRowsBelow(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertRowsBelow_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000219, nil, optArgs...)
	_= retVal
}

func (this *Selection) InsertColumnsRight()  {
	retVal, _ := this.Call(0x0000021a, nil)
	_= retVal
}

var Selection_InsertRowsAbove_OptArgs= []string{
	"NumRows", 
}

func (this *Selection) InsertRowsAbove(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertRowsAbove_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000021b, nil, optArgs...)
	_= retVal
}

func (this *Selection) RtlRun()  {
	retVal, _ := this.Call(0x00000258, nil)
	_= retVal
}

func (this *Selection) LtrRun()  {
	retVal, _ := this.Call(0x00000259, nil)
	_= retVal
}

func (this *Selection) BoldRun()  {
	retVal, _ := this.Call(0x0000025a, nil)
	_= retVal
}

func (this *Selection) ItalicRun()  {
	retVal, _ := this.Call(0x0000025b, nil)
	_= retVal
}

func (this *Selection) RtlPara()  {
	retVal, _ := this.Call(0x0000025d, nil)
	_= retVal
}

func (this *Selection) LtrPara()  {
	retVal, _ := this.Call(0x0000025e, nil)
	_= retVal
}

var Selection_InsertDateTime_OptArgs= []string{
	"DateTimeFormat", "InsertAsField", "InsertAsFullWidth", "DateLanguage", "CalendarType", 
}

func (this *Selection) InsertDateTime(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertDateTime_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bc, nil, optArgs...)
	_= retVal
}

var Selection_Sort2000_OptArgs= []string{
	"ExcludeHeader", "FieldNumber", "SortFieldType", "SortOrder", 
	"FieldNumber2", "SortFieldType2", "SortOrder2", "FieldNumber3", 
	"SortFieldType3", "SortOrder3", "SortColumn", "Separator", 
	"CaseSensitive", "BidiSort", "IgnoreThe", "IgnoreKashida", 
	"IgnoreDiacritics", "IgnoreHe", "LanguageID", 
}

func (this *Selection) Sort2000(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_Sort2000_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bd, nil, optArgs...)
	_= retVal
}

var Selection_ConvertToTable_OptArgs= []string{
	"Separator", "NumRows", "NumColumns", "InitialColumnWidth", 
	"Format", "ApplyBorders", "ApplyShading", "ApplyFont", 
	"ApplyColor", "ApplyHeadingRows", "ApplyLastRow", "ApplyFirstColumn", 
	"ApplyLastColumn", "AutoFit", "AutoFitBehavior", "DefaultTableBehavior", 
}

func (this *Selection) ConvertToTable(optArgs ...interface{}) *Table {
	optArgs = ole.ProcessOptArgs(Selection_ConvertToTable_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001c9, nil, optArgs...)
	return NewTable(retVal.IDispatch(), false, true)
}

func (this *Selection) NoProofing() int32 {
	retVal, _ := this.PropGet(0x000003ed, nil)
	return retVal.LValVal()
}

func (this *Selection) SetNoProofing(rhs int32)  {
	_ = this.PropPut(0x000003ed, []interface{}{rhs})
}

func (this *Selection) TopLevelTables() *Tables {
	retVal, _ := this.PropGet(0x000003ee, nil)
	return NewTables(retVal.IDispatch(), false, true)
}

func (this *Selection) LanguageDetected() bool {
	retVal, _ := this.PropGet(0x000003ef, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) SetLanguageDetected(rhs bool)  {
	_ = this.PropPut(0x000003ef, []interface{}{rhs})
}

func (this *Selection) FitTextWidth() float32 {
	retVal, _ := this.PropGet(0x000003f0, nil)
	return retVal.FltValVal()
}

func (this *Selection) SetFitTextWidth(rhs float32)  {
	_ = this.PropPut(0x000003f0, []interface{}{rhs})
}

func (this *Selection) ClearFormatting()  {
	retVal, _ := this.Call(0x000003f1, nil)
	_= retVal
}

func (this *Selection) PasteAppendTable()  {
	retVal, _ := this.Call(0x000003f2, nil)
	_= retVal
}

func (this *Selection) HTMLDivisions() *HTMLDivisions {
	retVal, _ := this.PropGet(0x000003f3, nil)
	return NewHTMLDivisions(retVal.IDispatch(), false, true)
}

func (this *Selection) SmartTags() *SmartTags {
	retVal, _ := this.PropGet(0x000003f7, nil)
	return NewSmartTags(retVal.IDispatch(), false, true)
}

func (this *Selection) ChildShapeRange() *ShapeRange {
	retVal, _ := this.PropGet(0x000003fd, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *Selection) HasChildShapeRange() bool {
	retVal, _ := this.PropGet(0x000003fe, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Selection) FootnoteOptions() *FootnoteOptions {
	retVal, _ := this.PropGet(0x00000400, nil)
	return NewFootnoteOptions(retVal.IDispatch(), false, true)
}

func (this *Selection) EndnoteOptions() *EndnoteOptions {
	retVal, _ := this.PropGet(0x00000401, nil)
	return NewEndnoteOptions(retVal.IDispatch(), false, true)
}

func (this *Selection) ToggleCharacterCode()  {
	retVal, _ := this.Call(0x000003f4, nil)
	_= retVal
}

func (this *Selection) PasteAndFormat(type_ int32)  {
	retVal, _ := this.Call(0x000003f5, []interface{}{type_})
	_= retVal
}

func (this *Selection) PasteExcelTable(linkedToExcel bool, wordFormatting bool, rtf bool)  {
	retVal, _ := this.Call(0x000003f6, []interface{}{linkedToExcel, wordFormatting, rtf})
	_= retVal
}

func (this *Selection) ShrinkDiscontiguousSelection()  {
	retVal, _ := this.Call(0x000003fb, nil)
	_= retVal
}

func (this *Selection) InsertStyleSeparator()  {
	retVal, _ := this.Call(0x000003fc, nil)
	_= retVal
}

var Selection_Sort_OptArgs= []string{
	"ExcludeHeader", "FieldNumber", "SortFieldType", "SortOrder", 
	"FieldNumber2", "SortFieldType2", "SortOrder2", "FieldNumber3", 
	"SortFieldType3", "SortOrder3", "SortColumn", "Separator", 
	"CaseSensitive", "BidiSort", "IgnoreThe", "IgnoreKashida", 
	"IgnoreDiacritics", "IgnoreHe", "LanguageID", "SubFieldNumber", 
	"SubFieldNumber2", "SubFieldNumber3", 
}

func (this *Selection) Sort(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_Sort_OptArgs, optArgs)
	retVal, _ := this.Call(0x000003ff, nil, optArgs...)
	_= retVal
}

func (this *Selection) XMLNodes() *XMLNodes {
	retVal, _ := this.PropGet(0x00000136, nil)
	return NewXMLNodes(retVal.IDispatch(), false, true)
}

func (this *Selection) XMLParentNode() *XMLNode {
	retVal, _ := this.PropGet(0x00000137, nil)
	return NewXMLNode(retVal.IDispatch(), false, true)
}

func (this *Selection) Editors() *Editors {
	retVal, _ := this.PropGet(0x00000139, nil)
	return NewEditors(retVal.IDispatch(), false, true)
}

var Selection_XML_OptArgs= []string{
	"DataOnly", 
}

func (this *Selection) XML(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Selection_XML_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000013a, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Selection) EnhMetaFileBits() ole.Variant {
	retVal, _ := this.PropGet(0x0000013b, nil)
	com.AddToScope(retVal)
	return *retVal
}

var Selection_GoToEditableRange_OptArgs= []string{
	"EditorID", 
}

func (this *Selection) GoToEditableRange(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Selection_GoToEditableRange_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000403, nil, optArgs...)
	return NewRange(retVal.IDispatch(), false, true)
}

var Selection_InsertXML_OptArgs= []string{
	"Transform", 
}

func (this *Selection) InsertXML(xml string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertXML_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000404, []interface{}{xml}, optArgs...)
	_= retVal
}

var Selection_InsertCaption_OptArgs= []string{
	"Title", "TitleAutoText", "Position", "ExcludeLabel", 
}

func (this *Selection) InsertCaption(label *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertCaption_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001a1, []interface{}{label}, optArgs...)
	_= retVal
}

var Selection_InsertCrossReference_OptArgs= []string{
	"InsertAsHyperlink", "IncludePosition", "SeparateNumbers", "SeparatorString", 
}

func (this *Selection) InsertCrossReference(referenceType *ole.Variant, referenceKind int32, referenceItem *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_InsertCrossReference_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001a2, []interface{}{referenceType, referenceKind, referenceItem}, optArgs...)
	_= retVal
}

func (this *Selection) OMaths() *OMaths {
	retVal, _ := this.PropGet(0x0000013c, nil)
	return NewOMaths(retVal.IDispatch(), false, true)
}

func (this *Selection) WordOpenXML() string {
	retVal, _ := this.PropGet(0x0000013d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Selection) ClearParagraphStyle()  {
	retVal, _ := this.Call(0x00000406, nil)
	_= retVal
}

func (this *Selection) ClearCharacterAllFormatting()  {
	retVal, _ := this.Call(0x00000407, nil)
	_= retVal
}

func (this *Selection) ClearCharacterStyle()  {
	retVal, _ := this.Call(0x00000408, nil)
	_= retVal
}

func (this *Selection) ClearCharacterDirectFormatting()  {
	retVal, _ := this.Call(0x00000409, nil)
	_= retVal
}

func (this *Selection) ContentControls() *ContentControls {
	retVal, _ := this.PropGet(0x0000040a, nil)
	return NewContentControls(retVal.IDispatch(), false, true)
}

func (this *Selection) ParentContentControl() *ContentControl {
	retVal, _ := this.PropGet(0x0000040b, nil)
	return NewContentControl(retVal.IDispatch(), false, true)
}

var Selection_ExportAsFixedFormat_OptArgs= []string{
	"OpenAfterExport", "OptimizeFor", "ExportCurrentPage", "Item", 
	"IncludeDocProps", "KeepIRM", "CreateBookmarks", "DocStructureTags", 
	"BitmapMissingFonts", "UseISO19005_1", "FixedFormatExtClassPtr", 
}

func (this *Selection) ExportAsFixedFormat(outputFileName string, exportFormat int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Selection_ExportAsFixedFormat_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000040c, []interface{}{outputFileName, exportFormat}, optArgs...)
	_= retVal
}

func (this *Selection) ReadingModeGrowFont()  {
	retVal, _ := this.Call(0x0000040d, nil)
	_= retVal
}

func (this *Selection) ReadingModeShrinkFont()  {
	retVal, _ := this.Call(0x0000040e, nil)
	_= retVal
}

func (this *Selection) ClearParagraphAllFormatting()  {
	retVal, _ := this.Call(0x0000040f, nil)
	_= retVal
}

func (this *Selection) ClearParagraphDirectFormatting()  {
	retVal, _ := this.Call(0x00000410, nil)
	_= retVal
}

func (this *Selection) InsertNewPage()  {
	retVal, _ := this.Call(0x00000411, nil)
	_= retVal
}

