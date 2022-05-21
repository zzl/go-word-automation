package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002095E-0000-0000-C000-000000000046
var IID_Range = syscall.GUID{0x0002095E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Range struct {
	ole.OleClient
}

func NewRange(pDisp *win32.IDispatch, addRef bool, scoped bool) *Range {
	p := &Range{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RangeFromVar(v ole.Variant) *Range {
	return NewRange(v.PdispValVal(), false, false)
}

func (this *Range) IID() *syscall.GUID {
	return &IID_Range
}

func (this *Range) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Range) Text() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Range) SetText(rhs string)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *Range) FormattedText() *Range {
	retVal := this.PropGet(0x00000002, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) SetFormattedText(rhs *Range)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *Range) Start() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Range) SetStart(rhs int32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *Range) End() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Range) SetEnd(rhs int32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *Range) Font() *Font {
	retVal := this.PropGet(0x00000005, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *Range) SetFont(rhs *Font)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *Range) Duplicate() *Range {
	retVal := this.PropGet(0x00000006, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) StoryType() int32 {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *Range) Tables() *Tables {
	retVal := this.PropGet(0x00000032, nil)
	return NewTables(retVal.PdispValVal(), false, true)
}

func (this *Range) Words() *Words {
	retVal := this.PropGet(0x00000033, nil)
	return NewWords(retVal.PdispValVal(), false, true)
}

func (this *Range) Sentences() *Sentences {
	retVal := this.PropGet(0x00000034, nil)
	return NewSentences(retVal.PdispValVal(), false, true)
}

func (this *Range) Characters() *Characters {
	retVal := this.PropGet(0x00000035, nil)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

func (this *Range) Footnotes() *Footnotes {
	retVal := this.PropGet(0x00000036, nil)
	return NewFootnotes(retVal.PdispValVal(), false, true)
}

func (this *Range) Endnotes() *Endnotes {
	retVal := this.PropGet(0x00000037, nil)
	return NewEndnotes(retVal.PdispValVal(), false, true)
}

func (this *Range) Comments() *Comments {
	retVal := this.PropGet(0x00000038, nil)
	return NewComments(retVal.PdispValVal(), false, true)
}

func (this *Range) Cells() *Cells {
	retVal := this.PropGet(0x00000039, nil)
	return NewCells(retVal.PdispValVal(), false, true)
}

func (this *Range) Sections() *Sections {
	retVal := this.PropGet(0x0000003a, nil)
	return NewSections(retVal.PdispValVal(), false, true)
}

func (this *Range) Paragraphs() *Paragraphs {
	retVal := this.PropGet(0x0000003b, nil)
	return NewParagraphs(retVal.PdispValVal(), false, true)
}

func (this *Range) Borders() *Borders {
	retVal := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *Range) SetBorders(rhs *Borders)  {
	retVal := this.PropPut(0x0000044c, []interface{}{rhs})
	_= retVal
}

func (this *Range) Shading() *Shading {
	retVal := this.PropGet(0x0000003d, nil)
	return NewShading(retVal.PdispValVal(), false, true)
}

func (this *Range) TextRetrievalMode() *TextRetrievalMode {
	retVal := this.PropGet(0x0000003e, nil)
	return NewTextRetrievalMode(retVal.PdispValVal(), false, true)
}

func (this *Range) SetTextRetrievalMode(rhs *TextRetrievalMode)  {
	retVal := this.PropPut(0x0000003e, []interface{}{rhs})
	_= retVal
}

func (this *Range) Fields() *Fields {
	retVal := this.PropGet(0x00000040, nil)
	return NewFields(retVal.PdispValVal(), false, true)
}

func (this *Range) FormFields() *FormFields {
	retVal := this.PropGet(0x00000041, nil)
	return NewFormFields(retVal.PdispValVal(), false, true)
}

func (this *Range) Frames() *Frames {
	retVal := this.PropGet(0x00000042, nil)
	return NewFrames(retVal.PdispValVal(), false, true)
}

func (this *Range) ParagraphFormat() *ParagraphFormat {
	retVal := this.PropGet(0x0000044e, nil)
	return NewParagraphFormat(retVal.PdispValVal(), false, true)
}

func (this *Range) SetParagraphFormat(rhs *ParagraphFormat)  {
	retVal := this.PropPut(0x0000044e, []interface{}{rhs})
	_= retVal
}

func (this *Range) ListFormat() *ListFormat {
	retVal := this.PropGet(0x00000044, nil)
	return NewListFormat(retVal.PdispValVal(), false, true)
}

func (this *Range) Bookmarks() *Bookmarks {
	retVal := this.PropGet(0x0000004b, nil)
	return NewBookmarks(retVal.PdispValVal(), false, true)
}

func (this *Range) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Range) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Range) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Range) Bold() int32 {
	retVal := this.PropGet(0x00000082, nil)
	return retVal.LValVal()
}

func (this *Range) SetBold(rhs int32)  {
	retVal := this.PropPut(0x00000082, []interface{}{rhs})
	_= retVal
}

func (this *Range) Italic() int32 {
	retVal := this.PropGet(0x00000083, nil)
	return retVal.LValVal()
}

func (this *Range) SetItalic(rhs int32)  {
	retVal := this.PropPut(0x00000083, []interface{}{rhs})
	_= retVal
}

func (this *Range) Underline() int32 {
	retVal := this.PropGet(0x0000008b, nil)
	return retVal.LValVal()
}

func (this *Range) SetUnderline(rhs int32)  {
	retVal := this.PropPut(0x0000008b, []interface{}{rhs})
	_= retVal
}

func (this *Range) EmphasisMark() int32 {
	retVal := this.PropGet(0x0000008c, nil)
	return retVal.LValVal()
}

func (this *Range) SetEmphasisMark(rhs int32)  {
	retVal := this.PropPut(0x0000008c, []interface{}{rhs})
	_= retVal
}

func (this *Range) DisableCharacterSpaceGrid() bool {
	retVal := this.PropGet(0x0000008d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) SetDisableCharacterSpaceGrid(rhs bool)  {
	retVal := this.PropPut(0x0000008d, []interface{}{rhs})
	_= retVal
}

func (this *Range) Revisions() *Revisions {
	retVal := this.PropGet(0x00000096, nil)
	return NewRevisions(retVal.PdispValVal(), false, true)
}

func (this *Range) Style() ole.Variant {
	retVal := this.PropGet(0x00000097, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) SetStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000097, []interface{}{rhs})
	_= retVal
}

func (this *Range) StoryLength() int32 {
	retVal := this.PropGet(0x00000098, nil)
	return retVal.LValVal()
}

func (this *Range) LanguageID() int32 {
	retVal := this.PropGet(0x00000099, nil)
	return retVal.LValVal()
}

func (this *Range) SetLanguageID(rhs int32)  {
	retVal := this.PropPut(0x00000099, []interface{}{rhs})
	_= retVal
}

func (this *Range) SynonymInfo() *SynonymInfo {
	retVal := this.PropGet(0x0000009b, nil)
	return NewSynonymInfo(retVal.PdispValVal(), false, true)
}

func (this *Range) Hyperlinks() *Hyperlinks {
	retVal := this.PropGet(0x0000009c, nil)
	return NewHyperlinks(retVal.PdispValVal(), false, true)
}

func (this *Range) ListParagraphs() *ListParagraphs {
	retVal := this.PropGet(0x0000009d, nil)
	return NewListParagraphs(retVal.PdispValVal(), false, true)
}

func (this *Range) Subdocuments() *Subdocuments {
	retVal := this.PropGet(0x0000009f, nil)
	return NewSubdocuments(retVal.PdispValVal(), false, true)
}

func (this *Range) GrammarChecked() bool {
	retVal := this.PropGet(0x00000104, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) SetGrammarChecked(rhs bool)  {
	retVal := this.PropPut(0x00000104, []interface{}{rhs})
	_= retVal
}

func (this *Range) SpellingChecked() bool {
	retVal := this.PropGet(0x00000105, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) SetSpellingChecked(rhs bool)  {
	retVal := this.PropPut(0x00000105, []interface{}{rhs})
	_= retVal
}

func (this *Range) HighlightColorIndex() int32 {
	retVal := this.PropGet(0x0000012d, nil)
	return retVal.LValVal()
}

func (this *Range) SetHighlightColorIndex(rhs int32)  {
	retVal := this.PropPut(0x0000012d, []interface{}{rhs})
	_= retVal
}

func (this *Range) Columns() *Columns {
	retVal := this.PropGet(0x0000012e, nil)
	return NewColumns(retVal.PdispValVal(), false, true)
}

func (this *Range) Rows() *Rows {
	retVal := this.PropGet(0x0000012f, nil)
	return NewRows(retVal.PdispValVal(), false, true)
}

func (this *Range) CanEdit() int32 {
	retVal := this.PropGet(0x00000130, nil)
	return retVal.LValVal()
}

func (this *Range) CanPaste() int32 {
	retVal := this.PropGet(0x00000131, nil)
	return retVal.LValVal()
}

func (this *Range) IsEndOfRowMark() bool {
	retVal := this.PropGet(0x00000133, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) BookmarkID() int32 {
	retVal := this.PropGet(0x00000134, nil)
	return retVal.LValVal()
}

func (this *Range) PreviousBookmarkID() int32 {
	retVal := this.PropGet(0x00000135, nil)
	return retVal.LValVal()
}

func (this *Range) Find() *Find {
	retVal := this.PropGet(0x00000106, nil)
	return NewFind(retVal.PdispValVal(), false, true)
}

func (this *Range) PageSetup() *PageSetup {
	retVal := this.PropGet(0x0000044d, nil)
	return NewPageSetup(retVal.PdispValVal(), false, true)
}

func (this *Range) SetPageSetup(rhs *PageSetup)  {
	retVal := this.PropPut(0x0000044d, []interface{}{rhs})
	_= retVal
}

func (this *Range) ShapeRange() *ShapeRange {
	retVal := this.PropGet(0x00000137, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *Range) Case() int32 {
	retVal := this.PropGet(0x00000138, nil)
	return retVal.LValVal()
}

func (this *Range) SetCase(rhs int32)  {
	retVal := this.PropPut(0x00000138, []interface{}{rhs})
	_= retVal
}

func (this *Range) Information(type_ int32) ole.Variant {
	retVal := this.PropGet(0x00000139, []interface{}{type_})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ReadabilityStatistics() *ReadabilityStatistics {
	retVal := this.PropGet(0x0000013a, nil)
	return NewReadabilityStatistics(retVal.PdispValVal(), false, true)
}

func (this *Range) GrammaticalErrors() *ProofreadingErrors {
	retVal := this.PropGet(0x0000013b, nil)
	return NewProofreadingErrors(retVal.PdispValVal(), false, true)
}

func (this *Range) SpellingErrors() *ProofreadingErrors {
	retVal := this.PropGet(0x0000013c, nil)
	return NewProofreadingErrors(retVal.PdispValVal(), false, true)
}

func (this *Range) Orientation() int32 {
	retVal := this.PropGet(0x0000013d, nil)
	return retVal.LValVal()
}

func (this *Range) SetOrientation(rhs int32)  {
	retVal := this.PropPut(0x0000013d, []interface{}{rhs})
	_= retVal
}

func (this *Range) InlineShapes() *InlineShapes {
	retVal := this.PropGet(0x0000013f, nil)
	return NewInlineShapes(retVal.PdispValVal(), false, true)
}

func (this *Range) NextStoryRange() *Range {
	retVal := this.PropGet(0x00000140, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) LanguageIDFarEast() int32 {
	retVal := this.PropGet(0x00000141, nil)
	return retVal.LValVal()
}

func (this *Range) SetLanguageIDFarEast(rhs int32)  {
	retVal := this.PropPut(0x00000141, []interface{}{rhs})
	_= retVal
}

func (this *Range) LanguageIDOther() int32 {
	retVal := this.PropGet(0x00000142, nil)
	return retVal.LValVal()
}

func (this *Range) SetLanguageIDOther(rhs int32)  {
	retVal := this.PropPut(0x00000142, []interface{}{rhs})
	_= retVal
}

func (this *Range) Select()  {
	retVal := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *Range) SetRange(start int32, end int32)  {
	retVal := this.Call(0x00000064, []interface{}{start, end})
	_= retVal
}

var Range_Collapse_OptArgs= []string{
	"Direction", 
}

func (this *Range) Collapse(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_Collapse_OptArgs, optArgs)
	retVal := this.Call(0x00000065, nil, optArgs...)
	_= retVal
}

func (this *Range) InsertBefore(text string)  {
	retVal := this.Call(0x00000066, []interface{}{text})
	_= retVal
}

func (this *Range) InsertAfter(text string)  {
	retVal := this.Call(0x00000068, []interface{}{text})
	_= retVal
}

var Range_Next_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Range) Next(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Next_OptArgs, optArgs)
	retVal := this.Call(0x00000069, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_Previous_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Range) Previous(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_Previous_OptArgs, optArgs)
	retVal := this.Call(0x0000006a, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_StartOf_OptArgs= []string{
	"Unit", "Extend", 
}

func (this *Range) StartOf(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_StartOf_OptArgs, optArgs)
	retVal := this.Call(0x0000006b, nil, optArgs...)
	return retVal.LValVal()
}

var Range_EndOf_OptArgs= []string{
	"Unit", "Extend", 
}

func (this *Range) EndOf(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_EndOf_OptArgs, optArgs)
	retVal := this.Call(0x0000006c, nil, optArgs...)
	return retVal.LValVal()
}

var Range_Move_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Range) Move(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_Move_OptArgs, optArgs)
	retVal := this.Call(0x0000006d, nil, optArgs...)
	return retVal.LValVal()
}

var Range_MoveStart_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Range) MoveStart(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_MoveStart_OptArgs, optArgs)
	retVal := this.Call(0x0000006e, nil, optArgs...)
	return retVal.LValVal()
}

var Range_MoveEnd_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Range) MoveEnd(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_MoveEnd_OptArgs, optArgs)
	retVal := this.Call(0x0000006f, nil, optArgs...)
	return retVal.LValVal()
}

var Range_MoveWhile_OptArgs= []string{
	"Count", 
}

func (this *Range) MoveWhile(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_MoveWhile_OptArgs, optArgs)
	retVal := this.Call(0x00000070, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Range_MoveStartWhile_OptArgs= []string{
	"Count", 
}

func (this *Range) MoveStartWhile(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_MoveStartWhile_OptArgs, optArgs)
	retVal := this.Call(0x00000071, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Range_MoveEndWhile_OptArgs= []string{
	"Count", 
}

func (this *Range) MoveEndWhile(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_MoveEndWhile_OptArgs, optArgs)
	retVal := this.Call(0x00000072, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Range_MoveUntil_OptArgs= []string{
	"Count", 
}

func (this *Range) MoveUntil(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_MoveUntil_OptArgs, optArgs)
	retVal := this.Call(0x00000073, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Range_MoveStartUntil_OptArgs= []string{
	"Count", 
}

func (this *Range) MoveStartUntil(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_MoveStartUntil_OptArgs, optArgs)
	retVal := this.Call(0x00000074, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

var Range_MoveEndUntil_OptArgs= []string{
	"Count", 
}

func (this *Range) MoveEndUntil(cset *ole.Variant, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_MoveEndUntil_OptArgs, optArgs)
	retVal := this.Call(0x00000075, []interface{}{cset}, optArgs...)
	return retVal.LValVal()
}

func (this *Range) Cut()  {
	retVal := this.Call(0x00000077, nil)
	_= retVal
}

func (this *Range) Copy()  {
	retVal := this.Call(0x00000078, nil)
	_= retVal
}

func (this *Range) Paste()  {
	retVal := this.Call(0x00000079, nil)
	_= retVal
}

var Range_InsertBreak_OptArgs= []string{
	"Type", 
}

func (this *Range) InsertBreak(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertBreak_OptArgs, optArgs)
	retVal := this.Call(0x0000007a, nil, optArgs...)
	_= retVal
}

var Range_InsertFile_OptArgs= []string{
	"Range", "ConfirmConversions", "Link", "Attachment", 
}

func (this *Range) InsertFile(fileName string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertFile_OptArgs, optArgs)
	retVal := this.Call(0x0000007b, []interface{}{fileName}, optArgs...)
	_= retVal
}

func (this *Range) InStory(range_ *Range) bool {
	retVal := this.Call(0x0000007d, []interface{}{range_})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) InRange(range_ *Range) bool {
	retVal := this.Call(0x0000007e, []interface{}{range_})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Range_Delete_OptArgs= []string{
	"Unit", "Count", 
}

func (this *Range) Delete(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_Delete_OptArgs, optArgs)
	retVal := this.Call(0x0000007f, nil, optArgs...)
	return retVal.LValVal()
}

func (this *Range) WholeStory()  {
	retVal := this.Call(0x00000080, nil)
	_= retVal
}

var Range_Expand_OptArgs= []string{
	"Unit", 
}

func (this *Range) Expand(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Range_Expand_OptArgs, optArgs)
	retVal := this.Call(0x00000081, nil, optArgs...)
	return retVal.LValVal()
}

func (this *Range) InsertParagraph()  {
	retVal := this.Call(0x000000a0, nil)
	_= retVal
}

func (this *Range) InsertParagraphAfter()  {
	retVal := this.Call(0x000000a1, nil)
	_= retVal
}

var Range_ConvertToTableOld_OptArgs= []string{
	"Separator", "NumRows", "NumColumns", "InitialColumnWidth", 
	"Format", "ApplyBorders", "ApplyShading", "ApplyFont", 
	"ApplyColor", "ApplyHeadingRows", "ApplyLastRow", "ApplyFirstColumn", 
	"ApplyLastColumn", "AutoFit", 
}

func (this *Range) ConvertToTableOld(optArgs ...interface{}) *Table {
	optArgs = ole.ProcessOptArgs(Range_ConvertToTableOld_OptArgs, optArgs)
	retVal := this.Call(0x000000a2, nil, optArgs...)
	return NewTable(retVal.PdispValVal(), false, true)
}

var Range_InsertDateTimeOld_OptArgs= []string{
	"DateTimeFormat", "InsertAsField", "InsertAsFullWidth", 
}

func (this *Range) InsertDateTimeOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertDateTimeOld_OptArgs, optArgs)
	retVal := this.Call(0x000000a3, nil, optArgs...)
	_= retVal
}

var Range_InsertSymbol_OptArgs= []string{
	"Font", "Unicode", "Bias", 
}

func (this *Range) InsertSymbol(characterNumber int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertSymbol_OptArgs, optArgs)
	retVal := this.Call(0x000000a4, []interface{}{characterNumber}, optArgs...)
	_= retVal
}

var Range_InsertCrossReference_2002_OptArgs= []string{
	"InsertAsHyperlink", "IncludePosition", 
}

func (this *Range) InsertCrossReference_2002(referenceType *ole.Variant, referenceKind int32, referenceItem *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertCrossReference_2002_OptArgs, optArgs)
	retVal := this.Call(0x000000a5, []interface{}{referenceType, referenceKind, referenceItem}, optArgs...)
	_= retVal
}

var Range_InsertCaptionXP_OptArgs= []string{
	"Title", "TitleAutoText", "Position", 
}

func (this *Range) InsertCaptionXP(label *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertCaptionXP_OptArgs, optArgs)
	retVal := this.Call(0x000000a6, []interface{}{label}, optArgs...)
	_= retVal
}

func (this *Range) CopyAsPicture()  {
	retVal := this.Call(0x000000a7, nil)
	_= retVal
}

var Range_SortOld_OptArgs= []string{
	"ExcludeHeader", "FieldNumber", "SortFieldType", "SortOrder", 
	"FieldNumber2", "SortFieldType2", "SortOrder2", "FieldNumber3", 
	"SortFieldType3", "SortOrder3", "SortColumn", "Separator", 
	"CaseSensitive", "LanguageID", 
}

func (this *Range) SortOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_SortOld_OptArgs, optArgs)
	retVal := this.Call(0x000000a8, nil, optArgs...)
	_= retVal
}

func (this *Range) SortAscending()  {
	retVal := this.Call(0x000000a9, nil)
	_= retVal
}

func (this *Range) SortDescending()  {
	retVal := this.Call(0x000000aa, nil)
	_= retVal
}

func (this *Range) IsEqual(range_ *Range) bool {
	retVal := this.Call(0x000000ab, []interface{}{range_})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) Calculate() float32 {
	retVal := this.Call(0x000000ac, nil)
	return retVal.FltValVal()
}

var Range_GoTo_OptArgs= []string{
	"What", "Which", "Count", "Name", 
}

func (this *Range) GoTo(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_GoTo_OptArgs, optArgs)
	retVal := this.Call(0x000000ad, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) GoToNext(what int32) *Range {
	retVal := this.Call(0x000000ae, []interface{}{what})
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Range) GoToPrevious(what int32) *Range {
	retVal := this.Call(0x000000af, []interface{}{what})
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_PasteSpecial_OptArgs= []string{
	"IconIndex", "Link", "Placement", "DisplayAsIcon", 
	"DataType", "IconFileName", "IconLabel", 
}

func (this *Range) PasteSpecial(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_PasteSpecial_OptArgs, optArgs)
	retVal := this.Call(0x000000b0, nil, optArgs...)
	_= retVal
}

func (this *Range) LookupNameProperties()  {
	retVal := this.Call(0x000000b1, nil)
	_= retVal
}

func (this *Range) ComputeStatistics(statistic int32) int32 {
	retVal := this.Call(0x000000b2, []interface{}{statistic})
	return retVal.LValVal()
}

func (this *Range) Relocate(direction int32)  {
	retVal := this.Call(0x000000b3, []interface{}{direction})
	_= retVal
}

func (this *Range) CheckSynonyms()  {
	retVal := this.Call(0x000000b4, nil)
	_= retVal
}

var Range_SubscribeTo_OptArgs= []string{
	"Format", 
}

func (this *Range) SubscribeTo(edition string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_SubscribeTo_OptArgs, optArgs)
	retVal := this.Call(0x000000b5, []interface{}{edition}, optArgs...)
	_= retVal
}

var Range_CreatePublisher_OptArgs= []string{
	"Edition", "ContainsPICT", "ContainsRTF", "ContainsText", 
}

func (this *Range) CreatePublisher(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_CreatePublisher_OptArgs, optArgs)
	retVal := this.Call(0x000000b6, nil, optArgs...)
	_= retVal
}

func (this *Range) InsertAutoText()  {
	retVal := this.Call(0x000000b7, nil)
	_= retVal
}

var Range_InsertDatabase_OptArgs= []string{
	"Format", "Style", "LinkToSource", "Connection", 
	"SQLStatement", "SQLStatement1", "PasswordDocument", "PasswordTemplate", 
	"WritePasswordDocument", "WritePasswordTemplate", "DataSource", "From", 
	"To", "IncludeFields", 
}

func (this *Range) InsertDatabase(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertDatabase_OptArgs, optArgs)
	retVal := this.Call(0x000000c2, nil, optArgs...)
	_= retVal
}

func (this *Range) AutoFormat()  {
	retVal := this.Call(0x000000c3, nil)
	_= retVal
}

func (this *Range) CheckGrammar()  {
	retVal := this.Call(0x000000cc, nil)
	_= retVal
}

var Range_CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "CustomDictionary2", 
	"CustomDictionary3", "CustomDictionary4", "CustomDictionary5", "CustomDictionary6", 
	"CustomDictionary7", "CustomDictionary8", "CustomDictionary9", "CustomDictionary10", 
}

func (this *Range) CheckSpelling(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x000000cd, nil, optArgs...)
	_= retVal
}

var Range_GetSpellingSuggestions_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "MainDictionary", "SuggestionMode", 
	"CustomDictionary2", "CustomDictionary3", "CustomDictionary4", "CustomDictionary5", 
	"CustomDictionary6", "CustomDictionary7", "CustomDictionary8", "CustomDictionary9", "CustomDictionary10", 
}

func (this *Range) GetSpellingSuggestions(optArgs ...interface{}) *SpellingSuggestions {
	optArgs = ole.ProcessOptArgs(Range_GetSpellingSuggestions_OptArgs, optArgs)
	retVal := this.Call(0x000000d1, nil, optArgs...)
	return NewSpellingSuggestions(retVal.PdispValVal(), false, true)
}

func (this *Range) InsertParagraphBefore()  {
	retVal := this.Call(0x000000d4, nil)
	_= retVal
}

func (this *Range) NextSubdocument()  {
	retVal := this.Call(0x000000db, nil)
	_= retVal
}

func (this *Range) PreviousSubdocument()  {
	retVal := this.Call(0x000000dc, nil)
	_= retVal
}

var Range_ConvertHangulAndHanja_OptArgs= []string{
	"ConversionsMode", "FastConversion", "CheckHangulEnding", "EnableRecentOrdering", "CustomDictionary", 
}

func (this *Range) ConvertHangulAndHanja(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_ConvertHangulAndHanja_OptArgs, optArgs)
	retVal := this.Call(0x000000dd, nil, optArgs...)
	_= retVal
}

func (this *Range) PasteAsNestedTable()  {
	retVal := this.Call(0x000000de, nil)
	_= retVal
}

var Range_ModifyEnclosure_OptArgs= []string{
	"Symbol", "EnclosedText", 
}

func (this *Range) ModifyEnclosure(style *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_ModifyEnclosure_OptArgs, optArgs)
	retVal := this.Call(0x000000df, []interface{}{style}, optArgs...)
	_= retVal
}

func (this *Range) PhoneticGuide(text string, alignment int32, raise int32, fontSize int32, fontName string)  {
	retVal := this.Call(0x000000e0, []interface{}{text, alignment, raise, fontSize, fontName})
	_= retVal
}

var Range_InsertDateTime_OptArgs= []string{
	"DateTimeFormat", "InsertAsField", "InsertAsFullWidth", "DateLanguage", "CalendarType", 
}

func (this *Range) InsertDateTime(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertDateTime_OptArgs, optArgs)
	retVal := this.Call(0x000001bc, nil, optArgs...)
	_= retVal
}

var Range_Sort_OptArgs= []string{
	"ExcludeHeader", "FieldNumber", "SortFieldType", "SortOrder", 
	"FieldNumber2", "SortFieldType2", "SortOrder2", "FieldNumber3", 
	"SortFieldType3", "SortOrder3", "SortColumn", "Separator", 
	"CaseSensitive", "BidiSort", "IgnoreThe", "IgnoreKashida", 
	"IgnoreDiacritics", "IgnoreHe", "LanguageID", 
}

func (this *Range) Sort(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_Sort_OptArgs, optArgs)
	retVal := this.Call(0x000001e4, nil, optArgs...)
	_= retVal
}

func (this *Range) DetectLanguage()  {
	retVal := this.Call(0x000000cb, nil)
	_= retVal
}

var Range_ConvertToTable_OptArgs= []string{
	"Separator", "NumRows", "NumColumns", "InitialColumnWidth", 
	"Format", "ApplyBorders", "ApplyShading", "ApplyFont", 
	"ApplyColor", "ApplyHeadingRows", "ApplyLastRow", "ApplyFirstColumn", 
	"ApplyLastColumn", "AutoFit", "AutoFitBehavior", "DefaultTableBehavior", 
}

func (this *Range) ConvertToTable(optArgs ...interface{}) *Table {
	optArgs = ole.ProcessOptArgs(Range_ConvertToTable_OptArgs, optArgs)
	retVal := this.Call(0x000001f2, nil, optArgs...)
	return NewTable(retVal.PdispValVal(), false, true)
}

func (this *Range) TCSCConverter(wdTCSCConverterDirection int32, commonTerms bool, useVariants bool)  {
	retVal := this.Call(0x000001f3, []interface{}{wdTCSCConverterDirection, commonTerms, useVariants})
	_= retVal
}

func (this *Range) LanguageDetected() bool {
	retVal := this.PropGet(0x00000107, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) SetLanguageDetected(rhs bool)  {
	retVal := this.PropPut(0x00000107, []interface{}{rhs})
	_= retVal
}

func (this *Range) FitTextWidth() float32 {
	retVal := this.PropGet(0x00000108, nil)
	return retVal.FltValVal()
}

func (this *Range) SetFitTextWidth(rhs float32)  {
	retVal := this.PropPut(0x00000108, []interface{}{rhs})
	_= retVal
}

func (this *Range) HorizontalInVertical() int32 {
	retVal := this.PropGet(0x00000109, nil)
	return retVal.LValVal()
}

func (this *Range) SetHorizontalInVertical(rhs int32)  {
	retVal := this.PropPut(0x00000109, []interface{}{rhs})
	_= retVal
}

func (this *Range) TwoLinesInOne() int32 {
	retVal := this.PropGet(0x0000010a, nil)
	return retVal.LValVal()
}

func (this *Range) SetTwoLinesInOne(rhs int32)  {
	retVal := this.PropPut(0x0000010a, []interface{}{rhs})
	_= retVal
}

func (this *Range) CombineCharacters() bool {
	retVal := this.PropGet(0x0000010b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) SetCombineCharacters(rhs bool)  {
	retVal := this.PropPut(0x0000010b, []interface{}{rhs})
	_= retVal
}

func (this *Range) NoProofing() int32 {
	retVal := this.PropGet(0x00000143, nil)
	return retVal.LValVal()
}

func (this *Range) SetNoProofing(rhs int32)  {
	retVal := this.PropPut(0x00000143, []interface{}{rhs})
	_= retVal
}

func (this *Range) TopLevelTables() *Tables {
	retVal := this.PropGet(0x00000144, nil)
	return NewTables(retVal.PdispValVal(), false, true)
}

func (this *Range) Scripts() *ole.DispatchClass {
	retVal := this.PropGet(0x00000145, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Range) CharacterWidth() int32 {
	retVal := this.PropGet(0x00000146, nil)
	return retVal.LValVal()
}

func (this *Range) SetCharacterWidth(rhs int32)  {
	retVal := this.PropPut(0x00000146, []interface{}{rhs})
	_= retVal
}

func (this *Range) Kana() int32 {
	retVal := this.PropGet(0x00000147, nil)
	return retVal.LValVal()
}

func (this *Range) SetKana(rhs int32)  {
	retVal := this.PropPut(0x00000147, []interface{}{rhs})
	_= retVal
}

func (this *Range) BoldBi() int32 {
	retVal := this.PropGet(0x00000190, nil)
	return retVal.LValVal()
}

func (this *Range) SetBoldBi(rhs int32)  {
	retVal := this.PropPut(0x00000190, []interface{}{rhs})
	_= retVal
}

func (this *Range) ItalicBi() int32 {
	retVal := this.PropGet(0x00000191, nil)
	return retVal.LValVal()
}

func (this *Range) SetItalicBi(rhs int32)  {
	retVal := this.PropPut(0x00000191, []interface{}{rhs})
	_= retVal
}

func (this *Range) ID() string {
	retVal := this.PropGet(0x00000195, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Range) SetID(rhs string)  {
	retVal := this.PropPut(0x00000195, []interface{}{rhs})
	_= retVal
}

func (this *Range) HTMLDivisions() *HTMLDivisions {
	retVal := this.PropGet(0x00000196, nil)
	return NewHTMLDivisions(retVal.PdispValVal(), false, true)
}

func (this *Range) SmartTags() *SmartTags {
	retVal := this.PropGet(0x00000197, nil)
	return NewSmartTags(retVal.PdispValVal(), false, true)
}

func (this *Range) ShowAll() bool {
	retVal := this.PropGet(0x00000198, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Range) SetShowAll(rhs bool)  {
	retVal := this.PropPut(0x00000198, []interface{}{rhs})
	_= retVal
}

func (this *Range) Document() *Document {
	retVal := this.PropGet(0x00000199, nil)
	return NewDocument(retVal.PdispValVal(), false, true)
}

func (this *Range) FootnoteOptions() *FootnoteOptions {
	retVal := this.PropGet(0x0000019a, nil)
	return NewFootnoteOptions(retVal.PdispValVal(), false, true)
}

func (this *Range) EndnoteOptions() *EndnoteOptions {
	retVal := this.PropGet(0x0000019b, nil)
	return NewEndnoteOptions(retVal.PdispValVal(), false, true)
}

func (this *Range) PasteAndFormat(type_ int32)  {
	retVal := this.Call(0x0000019c, []interface{}{type_})
	_= retVal
}

func (this *Range) PasteExcelTable(linkedToExcel bool, wordFormatting bool, rtf bool)  {
	retVal := this.Call(0x0000019d, []interface{}{linkedToExcel, wordFormatting, rtf})
	_= retVal
}

func (this *Range) PasteAppendTable()  {
	retVal := this.Call(0x0000019e, nil)
	_= retVal
}

func (this *Range) XMLNodes() *XMLNodes {
	retVal := this.PropGet(0x00000154, nil)
	return NewXMLNodes(retVal.PdispValVal(), false, true)
}

func (this *Range) XMLParentNode() *XMLNode {
	retVal := this.PropGet(0x00000155, nil)
	return NewXMLNode(retVal.PdispValVal(), false, true)
}

func (this *Range) Editors() *Editors {
	retVal := this.PropGet(0x00000157, nil)
	return NewEditors(retVal.PdispValVal(), false, true)
}

func (this *Range) XML(dataOnly bool) string {
	retVal := this.PropGet(0x00000158, []interface{}{dataOnly})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Range) EnhMetaFileBits() ole.Variant {
	retVal := this.PropGet(0x00000159, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

var Range_GoToEditableRange_OptArgs= []string{
	"EditorID", 
}

func (this *Range) GoToEditableRange(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Range_GoToEditableRange_OptArgs, optArgs)
	retVal := this.Call(0x0000019f, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Range_InsertXML_OptArgs= []string{
	"Transform", 
}

func (this *Range) InsertXML(xml string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertXML_OptArgs, optArgs)
	retVal := this.Call(0x000001a0, []interface{}{xml}, optArgs...)
	_= retVal
}

var Range_InsertCaption_OptArgs= []string{
	"Title", "TitleAutoText", "Position", "ExcludeLabel", 
}

func (this *Range) InsertCaption(label *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertCaption_OptArgs, optArgs)
	retVal := this.Call(0x000001a1, []interface{}{label}, optArgs...)
	_= retVal
}

var Range_InsertCrossReference_OptArgs= []string{
	"InsertAsHyperlink", "IncludePosition", "SeparateNumbers", "SeparatorString", 
}

func (this *Range) InsertCrossReference(referenceType *ole.Variant, referenceKind int32, referenceItem *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_InsertCrossReference_OptArgs, optArgs)
	retVal := this.Call(0x000001a2, []interface{}{referenceType, referenceKind, referenceItem}, optArgs...)
	_= retVal
}

func (this *Range) OMaths() *OMaths {
	retVal := this.PropGet(0x0000015a, nil)
	return NewOMaths(retVal.PdispValVal(), false, true)
}

func (this *Range) CharacterStyle() ole.Variant {
	retVal := this.PropGet(0x000001a4, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ParagraphStyle() ole.Variant {
	retVal := this.PropGet(0x000001a5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ListStyle() ole.Variant {
	retVal := this.PropGet(0x000001a6, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) TableStyle() ole.Variant {
	retVal := this.PropGet(0x000001a7, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Range) ContentControls() *ContentControls {
	retVal := this.PropGet(0x000001a8, nil)
	return NewContentControls(retVal.PdispValVal(), false, true)
}

func (this *Range) ExportFragment(fileName string, format int32)  {
	retVal := this.Call(0x000001a9, []interface{}{fileName, format})
	_= retVal
}

func (this *Range) WordOpenXML() string {
	retVal := this.PropGet(0x000001aa, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Range) SetListLevel(level int16)  {
	retVal := this.Call(0x000001ab, []interface{}{level})
	_= retVal
}

func (this *Range) InsertAlignmentTab(alignment int32, relativeTo int32)  {
	retVal := this.Call(0x000001f4, []interface{}{alignment, relativeTo})
	_= retVal
}

func (this *Range) ParentContentControl() *ContentControl {
	retVal := this.PropGet(0x000001f5, nil)
	return NewContentControl(retVal.PdispValVal(), false, true)
}

func (this *Range) ImportFragment(fileName string, matchDestination bool)  {
	retVal := this.Call(0x000001f6, []interface{}{fileName, matchDestination})
	_= retVal
}

var Range_ExportAsFixedFormat_OptArgs= []string{
	"FixedFormatExtClassPtr", 
}

func (this *Range) ExportAsFixedFormat(outputFileName string, exportFormat int32, openAfterExport bool, optimizeFor int32, exportCurrentPage bool, item int32, includeDocProps bool, keepIRM bool, createBookmarks int32, docStructureTags bool, bitmapMissingFonts bool, useISO19005_1 bool, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Range_ExportAsFixedFormat_OptArgs, optArgs)
	retVal := this.Call(0x000001f7, []interface{}{outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1}, optArgs...)
	_= retVal
}

func (this *Range) Locks() *CoAuthLocks {
	retVal := this.PropGet(0x000001f8, nil)
	return NewCoAuthLocks(retVal.PdispValVal(), false, true)
}

func (this *Range) Updates() *CoAuthUpdates {
	retVal := this.PropGet(0x000001f9, nil)
	return NewCoAuthUpdates(retVal.PdispValVal(), false, true)
}

func (this *Range) Conflicts() *Conflicts {
	retVal := this.PropGet(0x000001fa, nil)
	return NewConflicts(retVal.PdispValVal(), false, true)
}

