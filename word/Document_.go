package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002096B-0000-0000-C000-000000000046
var IID_Document_ = syscall.GUID{0x0002096B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Document_ struct {
	ole.OleClient
}

func NewDocument_(pDisp *win32.IDispatch, addRef bool, scoped bool) *Document_ {
	p := &Document_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Document_FromVar(v ole.Variant) *Document_ {
	return NewDocument_(v.PdispValVal(), false, false)
}

func (this *Document_) IID() *syscall.GUID {
	return &IID_Document_
}

func (this *Document_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Document_) Name() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) Application() *Application {
	retVal := this.PropGet(0x00000001, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Document_) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Document_) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) BuiltInDocumentProperties() *ole.DispatchClass {
	retVal := this.PropGet(0x000003e8, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) CustomDocumentProperties() *ole.DispatchClass {
	retVal := this.PropGet(0x00000002, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) Path() string {
	retVal := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) Bookmarks() *Bookmarks {
	retVal := this.PropGet(0x00000004, nil)
	return NewBookmarks(retVal.PdispValVal(), false, true)
}

func (this *Document_) Tables() *Tables {
	retVal := this.PropGet(0x00000006, nil)
	return NewTables(retVal.PdispValVal(), false, true)
}

func (this *Document_) Footnotes() *Footnotes {
	retVal := this.PropGet(0x00000007, nil)
	return NewFootnotes(retVal.PdispValVal(), false, true)
}

func (this *Document_) Endnotes() *Endnotes {
	retVal := this.PropGet(0x00000008, nil)
	return NewEndnotes(retVal.PdispValVal(), false, true)
}

func (this *Document_) Comments() *Comments {
	retVal := this.PropGet(0x00000009, nil)
	return NewComments(retVal.PdispValVal(), false, true)
}

func (this *Document_) Type() int32 {
	retVal := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *Document_) AutoHyphenation() bool {
	retVal := this.PropGet(0x0000000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetAutoHyphenation(rhs bool)  {
	retVal := this.PropPut(0x0000000b, []interface{}{rhs})
	_= retVal
}

func (this *Document_) HyphenateCaps() bool {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetHyphenateCaps(rhs bool)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *Document_) HyphenationZone() int32 {
	retVal := this.PropGet(0x0000000d, nil)
	return retVal.LValVal()
}

func (this *Document_) SetHyphenationZone(rhs int32)  {
	retVal := this.PropPut(0x0000000d, []interface{}{rhs})
	_= retVal
}

func (this *Document_) ConsecutiveHyphensLimit() int32 {
	retVal := this.PropGet(0x0000000e, nil)
	return retVal.LValVal()
}

func (this *Document_) SetConsecutiveHyphensLimit(rhs int32)  {
	retVal := this.PropPut(0x0000000e, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Sections() *Sections {
	retVal := this.PropGet(0x0000000f, nil)
	return NewSections(retVal.PdispValVal(), false, true)
}

func (this *Document_) Paragraphs() *Paragraphs {
	retVal := this.PropGet(0x00000010, nil)
	return NewParagraphs(retVal.PdispValVal(), false, true)
}

func (this *Document_) Words() *Words {
	retVal := this.PropGet(0x00000011, nil)
	return NewWords(retVal.PdispValVal(), false, true)
}

func (this *Document_) Sentences() *Sentences {
	retVal := this.PropGet(0x00000012, nil)
	return NewSentences(retVal.PdispValVal(), false, true)
}

func (this *Document_) Characters() *Characters {
	retVal := this.PropGet(0x00000013, nil)
	return NewCharacters(retVal.PdispValVal(), false, true)
}

func (this *Document_) Fields() *Fields {
	retVal := this.PropGet(0x00000014, nil)
	return NewFields(retVal.PdispValVal(), false, true)
}

func (this *Document_) FormFields() *FormFields {
	retVal := this.PropGet(0x00000015, nil)
	return NewFormFields(retVal.PdispValVal(), false, true)
}

func (this *Document_) Styles() *Styles {
	retVal := this.PropGet(0x00000016, nil)
	return NewStyles(retVal.PdispValVal(), false, true)
}

func (this *Document_) Frames() *Frames {
	retVal := this.PropGet(0x00000017, nil)
	return NewFrames(retVal.PdispValVal(), false, true)
}

func (this *Document_) TablesOfFigures() *TablesOfFigures {
	retVal := this.PropGet(0x00000019, nil)
	return NewTablesOfFigures(retVal.PdispValVal(), false, true)
}

func (this *Document_) Variables() *Variables {
	retVal := this.PropGet(0x0000001a, nil)
	return NewVariables(retVal.PdispValVal(), false, true)
}

func (this *Document_) MailMerge() *MailMerge {
	retVal := this.PropGet(0x0000001b, nil)
	return NewMailMerge(retVal.PdispValVal(), false, true)
}

func (this *Document_) Envelope() *Envelope {
	retVal := this.PropGet(0x0000001c, nil)
	return NewEnvelope(retVal.PdispValVal(), false, true)
}

func (this *Document_) FullName() string {
	retVal := this.PropGet(0x0000001d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) Revisions() *Revisions {
	retVal := this.PropGet(0x0000001e, nil)
	return NewRevisions(retVal.PdispValVal(), false, true)
}

func (this *Document_) TablesOfContents() *TablesOfContents {
	retVal := this.PropGet(0x0000001f, nil)
	return NewTablesOfContents(retVal.PdispValVal(), false, true)
}

func (this *Document_) TablesOfAuthorities() *TablesOfAuthorities {
	retVal := this.PropGet(0x00000020, nil)
	return NewTablesOfAuthorities(retVal.PdispValVal(), false, true)
}

func (this *Document_) PageSetup() *PageSetup {
	retVal := this.PropGet(0x0000044d, nil)
	return NewPageSetup(retVal.PdispValVal(), false, true)
}

func (this *Document_) SetPageSetup(rhs *PageSetup)  {
	retVal := this.PropPut(0x0000044d, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Windows() *Windows {
	retVal := this.PropGet(0x00000022, nil)
	return NewWindows(retVal.PdispValVal(), false, true)
}

func (this *Document_) HasRoutingSlip() bool {
	retVal := this.PropGet(0x00000023, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetHasRoutingSlip(rhs bool)  {
	retVal := this.PropPut(0x00000023, []interface{}{rhs})
	_= retVal
}

func (this *Document_) RoutingSlip() *RoutingSlip {
	retVal := this.PropGet(0x00000024, nil)
	return NewRoutingSlip(retVal.PdispValVal(), false, true)
}

func (this *Document_) Routed() bool {
	retVal := this.PropGet(0x00000025, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) TablesOfAuthoritiesCategories() *TablesOfAuthoritiesCategories {
	retVal := this.PropGet(0x00000026, nil)
	return NewTablesOfAuthoritiesCategories(retVal.PdispValVal(), false, true)
}

func (this *Document_) Indexes() *Indexes {
	retVal := this.PropGet(0x00000027, nil)
	return NewIndexes(retVal.PdispValVal(), false, true)
}

func (this *Document_) Saved() bool {
	retVal := this.PropGet(0x00000028, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetSaved(rhs bool)  {
	retVal := this.PropPut(0x00000028, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Content() *Range {
	retVal := this.PropGet(0x00000029, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Document_) ActiveWindow() *Window {
	retVal := this.PropGet(0x0000002a, nil)
	return NewWindow(retVal.PdispValVal(), false, true)
}

func (this *Document_) Kind() int32 {
	retVal := this.PropGet(0x0000002b, nil)
	return retVal.LValVal()
}

func (this *Document_) SetKind(rhs int32)  {
	retVal := this.PropPut(0x0000002b, []interface{}{rhs})
	_= retVal
}

func (this *Document_) ReadOnly() bool {
	retVal := this.PropGet(0x0000002c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) Subdocuments() *Subdocuments {
	retVal := this.PropGet(0x0000002d, nil)
	return NewSubdocuments(retVal.PdispValVal(), false, true)
}

func (this *Document_) IsMasterDocument() bool {
	retVal := this.PropGet(0x0000002e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) DefaultTabStop() float32 {
	retVal := this.PropGet(0x00000030, nil)
	return retVal.FltValVal()
}

func (this *Document_) SetDefaultTabStop(rhs float32)  {
	retVal := this.PropPut(0x00000030, []interface{}{rhs})
	_= retVal
}

func (this *Document_) EmbedTrueTypeFonts() bool {
	retVal := this.PropGet(0x00000032, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetEmbedTrueTypeFonts(rhs bool)  {
	retVal := this.PropPut(0x00000032, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SaveFormsData() bool {
	retVal := this.PropGet(0x00000033, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetSaveFormsData(rhs bool)  {
	retVal := this.PropPut(0x00000033, []interface{}{rhs})
	_= retVal
}

func (this *Document_) ReadOnlyRecommended() bool {
	retVal := this.PropGet(0x00000034, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetReadOnlyRecommended(rhs bool)  {
	retVal := this.PropPut(0x00000034, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SaveSubsetFonts() bool {
	retVal := this.PropGet(0x00000035, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetSaveSubsetFonts(rhs bool)  {
	retVal := this.PropPut(0x00000035, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Compatibility(type_ int32) bool {
	retVal := this.PropGet(0x00000037, []interface{}{type_})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetCompatibility(type_ int32, rhs bool)  {
	retVal := this.PropPut(0x00000037, []interface{}{type_, rhs})
	_= retVal
}

func (this *Document_) StoryRanges() *StoryRanges {
	retVal := this.PropGet(0x00000038, nil)
	return NewStoryRanges(retVal.PdispValVal(), false, true)
}

func (this *Document_) CommandBars() *ole.DispatchClass {
	retVal := this.PropGet(0x00000039, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) IsSubdocument() bool {
	retVal := this.PropGet(0x0000003a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SaveFormat() int32 {
	retVal := this.PropGet(0x0000003b, nil)
	return retVal.LValVal()
}

func (this *Document_) ProtectionType() int32 {
	retVal := this.PropGet(0x0000003c, nil)
	return retVal.LValVal()
}

func (this *Document_) Hyperlinks() *Hyperlinks {
	retVal := this.PropGet(0x0000003d, nil)
	return NewHyperlinks(retVal.PdispValVal(), false, true)
}

func (this *Document_) Shapes() *Shapes {
	retVal := this.PropGet(0x0000003e, nil)
	return NewShapes(retVal.PdispValVal(), false, true)
}

func (this *Document_) ListTemplates() *ListTemplates {
	retVal := this.PropGet(0x0000003f, nil)
	return NewListTemplates(retVal.PdispValVal(), false, true)
}

func (this *Document_) Lists() *Lists {
	retVal := this.PropGet(0x00000040, nil)
	return NewLists(retVal.PdispValVal(), false, true)
}

func (this *Document_) UpdateStylesOnOpen() bool {
	retVal := this.PropGet(0x00000042, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetUpdateStylesOnOpen(rhs bool)  {
	retVal := this.PropPut(0x00000042, []interface{}{rhs})
	_= retVal
}

func (this *Document_) AttachedTemplate() ole.Variant {
	retVal := this.PropGet(0x00000043, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Document_) SetAttachedTemplate(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000043, []interface{}{rhs})
	_= retVal
}

func (this *Document_) InlineShapes() *InlineShapes {
	retVal := this.PropGet(0x00000044, nil)
	return NewInlineShapes(retVal.PdispValVal(), false, true)
}

func (this *Document_) Background() *Shape {
	retVal := this.PropGet(0x00000045, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Document_) SetBackground(rhs *Shape)  {
	retVal := this.PropPut(0x00000045, []interface{}{rhs})
	_= retVal
}

func (this *Document_) GrammarChecked() bool {
	retVal := this.PropGet(0x00000046, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetGrammarChecked(rhs bool)  {
	retVal := this.PropPut(0x00000046, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SpellingChecked() bool {
	retVal := this.PropGet(0x00000047, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetSpellingChecked(rhs bool)  {
	retVal := this.PropPut(0x00000047, []interface{}{rhs})
	_= retVal
}

func (this *Document_) ShowGrammaticalErrors() bool {
	retVal := this.PropGet(0x00000048, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetShowGrammaticalErrors(rhs bool)  {
	retVal := this.PropPut(0x00000048, []interface{}{rhs})
	_= retVal
}

func (this *Document_) ShowSpellingErrors() bool {
	retVal := this.PropGet(0x00000049, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetShowSpellingErrors(rhs bool)  {
	retVal := this.PropPut(0x00000049, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Versions() *Versions {
	retVal := this.PropGet(0x0000004b, nil)
	return NewVersions(retVal.PdispValVal(), false, true)
}

func (this *Document_) ShowSummary() bool {
	retVal := this.PropGet(0x0000004c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetShowSummary(rhs bool)  {
	retVal := this.PropPut(0x0000004c, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SummaryViewMode() int32 {
	retVal := this.PropGet(0x0000004d, nil)
	return retVal.LValVal()
}

func (this *Document_) SetSummaryViewMode(rhs int32)  {
	retVal := this.PropPut(0x0000004d, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SummaryLength() int32 {
	retVal := this.PropGet(0x0000004e, nil)
	return retVal.LValVal()
}

func (this *Document_) SetSummaryLength(rhs int32)  {
	retVal := this.PropPut(0x0000004e, []interface{}{rhs})
	_= retVal
}

func (this *Document_) PrintFractionalWidths() bool {
	retVal := this.PropGet(0x0000004f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetPrintFractionalWidths(rhs bool)  {
	retVal := this.PropPut(0x0000004f, []interface{}{rhs})
	_= retVal
}

func (this *Document_) PrintPostScriptOverText() bool {
	retVal := this.PropGet(0x00000050, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetPrintPostScriptOverText(rhs bool)  {
	retVal := this.PropPut(0x00000050, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Container() *ole.DispatchClass {
	retVal := this.PropGet(0x00000052, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) PrintFormsData() bool {
	retVal := this.PropGet(0x00000053, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetPrintFormsData(rhs bool)  {
	retVal := this.PropPut(0x00000053, []interface{}{rhs})
	_= retVal
}

func (this *Document_) ListParagraphs() *ListParagraphs {
	retVal := this.PropGet(0x00000054, nil)
	return NewListParagraphs(retVal.PdispValVal(), false, true)
}

func (this *Document_) SetPassword(rhs string)  {
	retVal := this.PropPut(0x00000055, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SetWritePassword(rhs string)  {
	retVal := this.PropPut(0x00000056, []interface{}{rhs})
	_= retVal
}

func (this *Document_) HasPassword() bool {
	retVal := this.PropGet(0x00000057, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) WriteReserved() bool {
	retVal := this.PropGet(0x00000058, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) ActiveWritingStyle(languageID *ole.Variant) string {
	retVal := this.PropGet(0x0000005a, []interface{}{languageID})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) SetActiveWritingStyle(languageID *ole.Variant, rhs string)  {
	retVal := this.PropPut(0x0000005a, []interface{}{languageID, rhs})
	_= retVal
}

func (this *Document_) UserControl() bool {
	retVal := this.PropGet(0x0000005c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetUserControl(rhs bool)  {
	retVal := this.PropPut(0x0000005c, []interface{}{rhs})
	_= retVal
}

func (this *Document_) HasMailer() bool {
	retVal := this.PropGet(0x0000005d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetHasMailer(rhs bool)  {
	retVal := this.PropPut(0x0000005d, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Mailer() *Mailer {
	retVal := this.PropGet(0x0000005e, nil)
	return NewMailer(retVal.PdispValVal(), false, true)
}

func (this *Document_) ReadabilityStatistics() *ReadabilityStatistics {
	retVal := this.PropGet(0x00000060, nil)
	return NewReadabilityStatistics(retVal.PdispValVal(), false, true)
}

func (this *Document_) GrammaticalErrors() *ProofreadingErrors {
	retVal := this.PropGet(0x00000061, nil)
	return NewProofreadingErrors(retVal.PdispValVal(), false, true)
}

func (this *Document_) SpellingErrors() *ProofreadingErrors {
	retVal := this.PropGet(0x00000062, nil)
	return NewProofreadingErrors(retVal.PdispValVal(), false, true)
}

func (this *Document_) VBProject() *ole.DispatchClass {
	retVal := this.PropGet(0x00000063, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) FormsDesign() bool {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) CodeName_() string {
	retVal := this.PropGet(-2147418112, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) SetCodeName_(rhs string)  {
	retVal := this.PropPut(-2147418112, []interface{}{rhs})
	_= retVal
}

func (this *Document_) CodeName() string {
	retVal := this.PropGet(0x00000106, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) SnapToGrid() bool {
	retVal := this.PropGet(0x0000012c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetSnapToGrid(rhs bool)  {
	retVal := this.PropPut(0x0000012c, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SnapToShapes() bool {
	retVal := this.PropGet(0x0000012d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetSnapToShapes(rhs bool)  {
	retVal := this.PropPut(0x0000012d, []interface{}{rhs})
	_= retVal
}

func (this *Document_) GridDistanceHorizontal() float32 {
	retVal := this.PropGet(0x0000012e, nil)
	return retVal.FltValVal()
}

func (this *Document_) SetGridDistanceHorizontal(rhs float32)  {
	retVal := this.PropPut(0x0000012e, []interface{}{rhs})
	_= retVal
}

func (this *Document_) GridDistanceVertical() float32 {
	retVal := this.PropGet(0x0000012f, nil)
	return retVal.FltValVal()
}

func (this *Document_) SetGridDistanceVertical(rhs float32)  {
	retVal := this.PropPut(0x0000012f, []interface{}{rhs})
	_= retVal
}

func (this *Document_) GridOriginHorizontal() float32 {
	retVal := this.PropGet(0x00000130, nil)
	return retVal.FltValVal()
}

func (this *Document_) SetGridOriginHorizontal(rhs float32)  {
	retVal := this.PropPut(0x00000130, []interface{}{rhs})
	_= retVal
}

func (this *Document_) GridOriginVertical() float32 {
	retVal := this.PropGet(0x00000131, nil)
	return retVal.FltValVal()
}

func (this *Document_) SetGridOriginVertical(rhs float32)  {
	retVal := this.PropPut(0x00000131, []interface{}{rhs})
	_= retVal
}

func (this *Document_) GridSpaceBetweenHorizontalLines() int32 {
	retVal := this.PropGet(0x00000132, nil)
	return retVal.LValVal()
}

func (this *Document_) SetGridSpaceBetweenHorizontalLines(rhs int32)  {
	retVal := this.PropPut(0x00000132, []interface{}{rhs})
	_= retVal
}

func (this *Document_) GridSpaceBetweenVerticalLines() int32 {
	retVal := this.PropGet(0x00000133, nil)
	return retVal.LValVal()
}

func (this *Document_) SetGridSpaceBetweenVerticalLines(rhs int32)  {
	retVal := this.PropPut(0x00000133, []interface{}{rhs})
	_= retVal
}

func (this *Document_) GridOriginFromMargin() bool {
	retVal := this.PropGet(0x00000134, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetGridOriginFromMargin(rhs bool)  {
	retVal := this.PropPut(0x00000134, []interface{}{rhs})
	_= retVal
}

func (this *Document_) KerningByAlgorithm() bool {
	retVal := this.PropGet(0x00000135, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetKerningByAlgorithm(rhs bool)  {
	retVal := this.PropPut(0x00000135, []interface{}{rhs})
	_= retVal
}

func (this *Document_) JustificationMode() int32 {
	retVal := this.PropGet(0x00000136, nil)
	return retVal.LValVal()
}

func (this *Document_) SetJustificationMode(rhs int32)  {
	retVal := this.PropPut(0x00000136, []interface{}{rhs})
	_= retVal
}

func (this *Document_) FarEastLineBreakLevel() int32 {
	retVal := this.PropGet(0x00000137, nil)
	return retVal.LValVal()
}

func (this *Document_) SetFarEastLineBreakLevel(rhs int32)  {
	retVal := this.PropPut(0x00000137, []interface{}{rhs})
	_= retVal
}

func (this *Document_) NoLineBreakBefore() string {
	retVal := this.PropGet(0x00000138, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) SetNoLineBreakBefore(rhs string)  {
	retVal := this.PropPut(0x00000138, []interface{}{rhs})
	_= retVal
}

func (this *Document_) NoLineBreakAfter() string {
	retVal := this.PropGet(0x00000139, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) SetNoLineBreakAfter(rhs string)  {
	retVal := this.PropPut(0x00000139, []interface{}{rhs})
	_= retVal
}

func (this *Document_) TrackRevisions() bool {
	retVal := this.PropGet(0x0000013a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetTrackRevisions(rhs bool)  {
	retVal := this.PropPut(0x0000013a, []interface{}{rhs})
	_= retVal
}

func (this *Document_) PrintRevisions() bool {
	retVal := this.PropGet(0x0000013b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetPrintRevisions(rhs bool)  {
	retVal := this.PropPut(0x0000013b, []interface{}{rhs})
	_= retVal
}

func (this *Document_) ShowRevisions() bool {
	retVal := this.PropGet(0x0000013c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetShowRevisions(rhs bool)  {
	retVal := this.PropPut(0x0000013c, []interface{}{rhs})
	_= retVal
}

var Document__Close_OptArgs= []string{
	"SaveChanges", "OriginalFormat", "RouteDocument", 
}

func (this *Document_) Close(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__Close_OptArgs, optArgs)
	retVal := this.Call(0x00000451, nil, optArgs...)
	_= retVal
}

var Document__SaveAs2000_OptArgs= []string{
	"FileName", "FileFormat", "LockComments", "Password", 
	"AddToRecentFiles", "WritePassword", "ReadOnlyRecommended", "EmbedTrueTypeFonts", 
	"SaveNativePictureFormat", "SaveFormsData", "SaveAsAOCELetter", 
}

func (this *Document_) SaveAs2000(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__SaveAs2000_OptArgs, optArgs)
	retVal := this.Call(0x00000066, nil, optArgs...)
	_= retVal
}

func (this *Document_) Repaginate()  {
	retVal := this.Call(0x00000067, nil)
	_= retVal
}

func (this *Document_) FitToPages()  {
	retVal := this.Call(0x00000068, nil)
	_= retVal
}

func (this *Document_) ManualHyphenation()  {
	retVal := this.Call(0x00000069, nil)
	_= retVal
}

func (this *Document_) Select()  {
	retVal := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *Document_) DataForm()  {
	retVal := this.Call(0x0000006a, nil)
	_= retVal
}

func (this *Document_) Route()  {
	retVal := this.Call(0x0000006b, nil)
	_= retVal
}

func (this *Document_) Save()  {
	retVal := this.Call(0x0000006c, nil)
	_= retVal
}

var Document__PrintOutOld_OptArgs= []string{
	"Background", "Append", "Range", "OutputFileName", 
	"From", "To", "Item", "Copies", 
	"Pages", "PageType", "PrintToFile", "Collate", 
	"ActivePrinterMacGX", "ManualDuplexPrint", 
}

func (this *Document_) PrintOutOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__PrintOutOld_OptArgs, optArgs)
	retVal := this.Call(0x0000006d, nil, optArgs...)
	_= retVal
}

func (this *Document_) SendMail()  {
	retVal := this.Call(0x0000006e, nil)
	_= retVal
}

var Document__Range_OptArgs= []string{
	"Start", "End", 
}

func (this *Document_) Range(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Document__Range_OptArgs, optArgs)
	retVal := this.Call(0x000007d0, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Document_) RunAutoMacro(which int32)  {
	retVal := this.Call(0x00000070, []interface{}{which})
	_= retVal
}

func (this *Document_) Activate()  {
	retVal := this.Call(0x00000071, nil)
	_= retVal
}

func (this *Document_) PrintPreview()  {
	retVal := this.Call(0x00000072, nil)
	_= retVal
}

var Document__GoTo_OptArgs= []string{
	"What", "Which", "Count", "Name", 
}

func (this *Document_) GoTo(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Document__GoTo_OptArgs, optArgs)
	retVal := this.Call(0x00000073, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Document__Undo_OptArgs= []string{
	"Times", 
}

func (this *Document_) Undo(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Document__Undo_OptArgs, optArgs)
	retVal := this.Call(0x00000074, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Document__Redo_OptArgs= []string{
	"Times", 
}

func (this *Document_) Redo(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Document__Redo_OptArgs, optArgs)
	retVal := this.Call(0x00000075, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Document__ComputeStatistics_OptArgs= []string{
	"IncludeFootnotesAndEndnotes", 
}

func (this *Document_) ComputeStatistics(statistic int32, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Document__ComputeStatistics_OptArgs, optArgs)
	retVal := this.Call(0x00000076, []interface{}{statistic}, optArgs...)
	return retVal.LValVal()
}

func (this *Document_) MakeCompatibilityDefault()  {
	retVal := this.Call(0x00000077, nil)
	_= retVal
}

var Document__Protect2002_OptArgs= []string{
	"NoReset", "Password", 
}

func (this *Document_) Protect2002(type_ int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__Protect2002_OptArgs, optArgs)
	retVal := this.Call(0x00000078, []interface{}{type_}, optArgs...)
	_= retVal
}

var Document__Unprotect_OptArgs= []string{
	"Password", 
}

func (this *Document_) Unprotect(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__Unprotect_OptArgs, optArgs)
	retVal := this.Call(0x00000079, nil, optArgs...)
	_= retVal
}

var Document__EditionOptions_OptArgs= []string{
	"Format", 
}

func (this *Document_) EditionOptions(type_ int32, option int32, name string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__EditionOptions_OptArgs, optArgs)
	retVal := this.Call(0x0000007a, []interface{}{type_, option, name}, optArgs...)
	_= retVal
}

var Document__RunLetterWizard_OptArgs= []string{
	"LetterContent", "WizardMode", 
}

func (this *Document_) RunLetterWizard(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__RunLetterWizard_OptArgs, optArgs)
	retVal := this.Call(0x0000007b, nil, optArgs...)
	_= retVal
}

func (this *Document_) GetLetterContent() *LetterContent {
	retVal := this.Call(0x0000007c, nil)
	return NewLetterContent(retVal.PdispValVal(), false, true)
}

func (this *Document_) SetLetterContent(letterContent *ole.Variant)  {
	retVal := this.Call(0x0000007d, []interface{}{letterContent})
	_= retVal
}

func (this *Document_) CopyStylesFromTemplate(template string)  {
	retVal := this.Call(0x0000007e, []interface{}{template})
	_= retVal
}

func (this *Document_) UpdateStyles()  {
	retVal := this.Call(0x0000007f, nil)
	_= retVal
}

func (this *Document_) CheckGrammar()  {
	retVal := this.Call(0x00000083, nil)
	_= retVal
}

var Document__CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "AlwaysSuggest", "CustomDictionary2", 
	"CustomDictionary3", "CustomDictionary4", "CustomDictionary5", "CustomDictionary6", 
	"CustomDictionary7", "CustomDictionary8", "CustomDictionary9", "CustomDictionary10", 
}

func (this *Document_) CheckSpelling(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__CheckSpelling_OptArgs, optArgs)
	retVal := this.Call(0x00000084, nil, optArgs...)
	_= retVal
}

var Document__FollowHyperlink_OptArgs= []string{
	"Address", "SubAddress", "NewWindow", "AddHistory", 
	"ExtraInfo", "Method", "HeaderInfo", 
}

func (this *Document_) FollowHyperlink(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__FollowHyperlink_OptArgs, optArgs)
	retVal := this.Call(0x00000087, nil, optArgs...)
	_= retVal
}

func (this *Document_) AddToFavorites()  {
	retVal := this.Call(0x00000088, nil)
	_= retVal
}

func (this *Document_) Reload()  {
	retVal := this.Call(0x00000089, nil)
	_= retVal
}

var Document__AutoSummarize_OptArgs= []string{
	"Length", "Mode", "UpdateProperties", 
}

func (this *Document_) AutoSummarize(optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(Document__AutoSummarize_OptArgs, optArgs)
	retVal := this.Call(0x0000008a, nil, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

var Document__RemoveNumbers_OptArgs= []string{
	"NumberType", 
}

func (this *Document_) RemoveNumbers(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__RemoveNumbers_OptArgs, optArgs)
	retVal := this.Call(0x0000008c, nil, optArgs...)
	_= retVal
}

var Document__ConvertNumbersToText_OptArgs= []string{
	"NumberType", 
}

func (this *Document_) ConvertNumbersToText(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__ConvertNumbersToText_OptArgs, optArgs)
	retVal := this.Call(0x0000008d, nil, optArgs...)
	_= retVal
}

var Document__CountNumberedItems_OptArgs= []string{
	"NumberType", "Level", 
}

func (this *Document_) CountNumberedItems(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Document__CountNumberedItems_OptArgs, optArgs)
	retVal := this.Call(0x0000008e, nil, optArgs...)
	return retVal.LValVal()
}

func (this *Document_) Post()  {
	retVal := this.Call(0x0000008f, nil)
	_= retVal
}

func (this *Document_) ToggleFormsDesign()  {
	retVal := this.Call(0x00000090, nil)
	_= retVal
}

func (this *Document_) Compare2000(name string)  {
	retVal := this.Call(0x00000091, []interface{}{name})
	_= retVal
}

func (this *Document_) UpdateSummaryProperties()  {
	retVal := this.Call(0x00000092, nil)
	_= retVal
}

func (this *Document_) GetCrossReferenceItems(referenceType *ole.Variant) ole.Variant {
	retVal := this.Call(0x00000093, []interface{}{referenceType})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Document_) AutoFormat()  {
	retVal := this.Call(0x00000094, nil)
	_= retVal
}

func (this *Document_) ViewCode()  {
	retVal := this.Call(0x00000095, nil)
	_= retVal
}

func (this *Document_) ViewPropertyBrowser()  {
	retVal := this.Call(0x00000096, nil)
	_= retVal
}

func (this *Document_) ForwardMailer()  {
	retVal := this.Call(0x000000fa, nil)
	_= retVal
}

func (this *Document_) Reply()  {
	retVal := this.Call(0x000000fb, nil)
	_= retVal
}

func (this *Document_) ReplyAll()  {
	retVal := this.Call(0x000000fc, nil)
	_= retVal
}

var Document__SendMailer_OptArgs= []string{
	"FileFormat", "Priority", 
}

func (this *Document_) SendMailer(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__SendMailer_OptArgs, optArgs)
	retVal := this.Call(0x000000fd, nil, optArgs...)
	_= retVal
}

func (this *Document_) UndoClear()  {
	retVal := this.Call(0x000000fe, nil)
	_= retVal
}

func (this *Document_) PresentIt()  {
	retVal := this.Call(0x000000ff, nil)
	_= retVal
}

var Document__SendFax_OptArgs= []string{
	"Subject", 
}

func (this *Document_) SendFax(address string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__SendFax_OptArgs, optArgs)
	retVal := this.Call(0x00000100, []interface{}{address}, optArgs...)
	_= retVal
}

func (this *Document_) Merge2000(fileName string)  {
	retVal := this.Call(0x00000101, []interface{}{fileName})
	_= retVal
}

func (this *Document_) ClosePrintPreview()  {
	retVal := this.Call(0x00000102, nil)
	_= retVal
}

func (this *Document_) CheckConsistency()  {
	retVal := this.Call(0x00000103, nil)
	_= retVal
}

var Document__CreateLetterContent_OptArgs= []string{
	"InfoBlock", "RecipientCode", "RecipientGender", "ReturnAddressShortForm", 
	"SenderCity", "SenderCode", "SenderGender", "SenderReference", 
}

func (this *Document_) CreateLetterContent(dateFormat string, includeHeaderFooter bool, pageDesign string, letterStyle int32, letterhead bool, letterheadLocation int32, letterheadSize float32, recipientName string, recipientAddress string, salutation string, salutationType int32, recipientReference string, mailingInstructions string, attentionLine string, subject string, cclist string, returnAddress string, senderName string, closing string, senderCompany string, senderJobTitle string, senderInitials string, enclosureNumber int32, optArgs ...interface{}) *LetterContent {
	optArgs = ole.ProcessOptArgs(Document__CreateLetterContent_OptArgs, optArgs)
	retVal := this.Call(0x00000104, []interface{}{dateFormat, includeHeaderFooter, pageDesign, letterStyle, letterhead, letterheadLocation, letterheadSize, recipientName, recipientAddress, salutation, salutationType, recipientReference, mailingInstructions, attentionLine, subject, cclist, returnAddress, senderName, closing, senderCompany, senderJobTitle, senderInitials, enclosureNumber}, optArgs...)
	return NewLetterContent(retVal.PdispValVal(), false, true)
}

func (this *Document_) AcceptAllRevisions()  {
	retVal := this.Call(0x0000013d, nil)
	_= retVal
}

func (this *Document_) RejectAllRevisions()  {
	retVal := this.Call(0x0000013e, nil)
	_= retVal
}

func (this *Document_) DetectLanguage()  {
	retVal := this.Call(0x00000097, nil)
	_= retVal
}

func (this *Document_) ApplyTheme(name string)  {
	retVal := this.Call(0x00000142, []interface{}{name})
	_= retVal
}

func (this *Document_) RemoveTheme()  {
	retVal := this.Call(0x00000143, nil)
	_= retVal
}

func (this *Document_) WebPagePreview()  {
	retVal := this.Call(0x00000145, nil)
	_= retVal
}

func (this *Document_) ReloadAs(encoding int32)  {
	retVal := this.Call(0x0000014b, []interface{}{encoding})
	_= retVal
}

func (this *Document_) ActiveTheme() string {
	retVal := this.PropGet(0x0000021c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) ActiveThemeDisplayName() string {
	retVal := this.PropGet(0x0000021d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) Email() *Email {
	retVal := this.PropGet(0x0000013f, nil)
	return NewEmail(retVal.PdispValVal(), false, true)
}

func (this *Document_) Scripts() *ole.DispatchClass {
	retVal := this.PropGet(0x00000140, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) LanguageDetected() bool {
	retVal := this.PropGet(0x00000141, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetLanguageDetected(rhs bool)  {
	retVal := this.PropPut(0x00000141, []interface{}{rhs})
	_= retVal
}

func (this *Document_) FarEastLineBreakLanguage() int32 {
	retVal := this.PropGet(0x00000146, nil)
	return retVal.LValVal()
}

func (this *Document_) SetFarEastLineBreakLanguage(rhs int32)  {
	retVal := this.PropPut(0x00000146, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Frameset() *Frameset {
	retVal := this.PropGet(0x00000147, nil)
	return NewFrameset(retVal.PdispValVal(), false, true)
}

func (this *Document_) ClickAndTypeParagraphStyle() ole.Variant {
	retVal := this.PropGet(0x00000148, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Document_) SetClickAndTypeParagraphStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000148, []interface{}{rhs})
	_= retVal
}

func (this *Document_) HTMLProject() *ole.DispatchClass {
	retVal := this.PropGet(0x00000149, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) WebOptions() *WebOptions {
	retVal := this.PropGet(0x0000014a, nil)
	return NewWebOptions(retVal.PdispValVal(), false, true)
}

func (this *Document_) OpenEncoding() int32 {
	retVal := this.PropGet(0x0000014c, nil)
	return retVal.LValVal()
}

func (this *Document_) SaveEncoding() int32 {
	retVal := this.PropGet(0x0000014d, nil)
	return retVal.LValVal()
}

func (this *Document_) SetSaveEncoding(rhs int32)  {
	retVal := this.PropPut(0x0000014d, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OptimizeForWord97() bool {
	retVal := this.PropGet(0x0000014e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetOptimizeForWord97(rhs bool)  {
	retVal := this.PropPut(0x0000014e, []interface{}{rhs})
	_= retVal
}

func (this *Document_) VBASigned() bool {
	retVal := this.PropGet(0x0000014f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Document__PrintOut2000_OptArgs= []string{
	"Background", "Append", "Range", "OutputFileName", 
	"From", "To", "Item", "Copies", 
	"Pages", "PageType", "PrintToFile", "Collate", 
	"ActivePrinterMacGX", "ManualDuplexPrint", "PrintZoomColumn", "PrintZoomRow", 
	"PrintZoomPaperWidth", "PrintZoomPaperHeight", 
}

func (this *Document_) PrintOut2000(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__PrintOut2000_OptArgs, optArgs)
	retVal := this.Call(0x000001bc, nil, optArgs...)
	_= retVal
}

func (this *Document_) Sblt(s string)  {
	retVal := this.Call(0x000001bd, []interface{}{s})
	_= retVal
}

func (this *Document_) ConvertVietDoc(codePageOrigin int32)  {
	retVal := this.Call(0x000001bf, []interface{}{codePageOrigin})
	_= retVal
}

var Document__PrintOut_OptArgs= []string{
	"Background", "Append", "Range", "OutputFileName", 
	"From", "To", "Item", "Copies", 
	"Pages", "PageType", "PrintToFile", "Collate", 
	"ActivePrinterMacGX", "ManualDuplexPrint", "PrintZoomColumn", "PrintZoomRow", 
	"PrintZoomPaperWidth", "PrintZoomPaperHeight", 
}

func (this *Document_) PrintOut(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__PrintOut_OptArgs, optArgs)
	retVal := this.Call(0x000001be, nil, optArgs...)
	_= retVal
}

func (this *Document_) MailEnvelope() *ole.DispatchClass {
	retVal := this.PropGet(0x00000150, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) DisableFeatures() bool {
	retVal := this.PropGet(0x00000151, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetDisableFeatures(rhs bool)  {
	retVal := this.PropPut(0x00000151, []interface{}{rhs})
	_= retVal
}

func (this *Document_) DoNotEmbedSystemFonts() bool {
	retVal := this.PropGet(0x00000152, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetDoNotEmbedSystemFonts(rhs bool)  {
	retVal := this.PropPut(0x00000152, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Signatures() *ole.DispatchClass {
	retVal := this.PropGet(0x00000153, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) DefaultTargetFrame() string {
	retVal := this.PropGet(0x00000154, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) SetDefaultTargetFrame(rhs string)  {
	retVal := this.PropPut(0x00000154, []interface{}{rhs})
	_= retVal
}

func (this *Document_) HTMLDivisions() *HTMLDivisions {
	retVal := this.PropGet(0x00000156, nil)
	return NewHTMLDivisions(retVal.PdispValVal(), false, true)
}

func (this *Document_) DisableFeaturesIntroducedAfter() int32 {
	retVal := this.PropGet(0x00000157, nil)
	return retVal.LValVal()
}

func (this *Document_) SetDisableFeaturesIntroducedAfter(rhs int32)  {
	retVal := this.PropPut(0x00000157, []interface{}{rhs})
	_= retVal
}

func (this *Document_) RemovePersonalInformation() bool {
	retVal := this.PropGet(0x00000158, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetRemovePersonalInformation(rhs bool)  {
	retVal := this.PropPut(0x00000158, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SmartTags() *SmartTags {
	retVal := this.PropGet(0x0000015a, nil)
	return NewSmartTags(retVal.PdispValVal(), false, true)
}

var Document__Compare2002_OptArgs= []string{
	"AuthorName", "CompareTarget", "DetectFormatChanges", "IgnoreAllComparisonWarnings", "AddToRecentFiles", 
}

func (this *Document_) Compare2002(name string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__Compare2002_OptArgs, optArgs)
	retVal := this.Call(0x00000159, []interface{}{name}, optArgs...)
	_= retVal
}

func (this *Document_) CheckIn(saveChanges bool, comments *ole.Variant, makePublic bool)  {
	retVal := this.Call(0x0000015d, []interface{}{saveChanges, comments, makePublic})
	_= retVal
}

func (this *Document_) CanCheckin() bool {
	retVal := this.Call(0x0000015f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Document__Merge_OptArgs= []string{
	"MergeTarget", "DetectFormatChanges", "UseFormattingFrom", "AddToRecentFiles", 
}

func (this *Document_) Merge(fileName string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__Merge_OptArgs, optArgs)
	retVal := this.Call(0x0000016a, []interface{}{fileName}, optArgs...)
	_= retVal
}

func (this *Document_) EmbedSmartTags() bool {
	retVal := this.PropGet(0x0000015b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetEmbedSmartTags(rhs bool)  {
	retVal := this.PropPut(0x0000015b, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SmartTagsAsXMLProps() bool {
	retVal := this.PropGet(0x0000015c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetSmartTagsAsXMLProps(rhs bool)  {
	retVal := this.PropPut(0x0000015c, []interface{}{rhs})
	_= retVal
}

func (this *Document_) TextEncoding() int32 {
	retVal := this.PropGet(0x00000165, nil)
	return retVal.LValVal()
}

func (this *Document_) SetTextEncoding(rhs int32)  {
	retVal := this.PropPut(0x00000165, []interface{}{rhs})
	_= retVal
}

func (this *Document_) TextLineEnding() int32 {
	retVal := this.PropGet(0x00000166, nil)
	return retVal.LValVal()
}

func (this *Document_) SetTextLineEnding(rhs int32)  {
	retVal := this.PropPut(0x00000166, []interface{}{rhs})
	_= retVal
}

var Document__SendForReview_OptArgs= []string{
	"Recipients", "Subject", "ShowMessage", "IncludeAttachment", 
}

func (this *Document_) SendForReview(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__SendForReview_OptArgs, optArgs)
	retVal := this.Call(0x00000161, nil, optArgs...)
	_= retVal
}

var Document__ReplyWithChanges_OptArgs= []string{
	"ShowMessage", 
}

func (this *Document_) ReplyWithChanges(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__ReplyWithChanges_OptArgs, optArgs)
	retVal := this.Call(0x00000162, nil, optArgs...)
	_= retVal
}

func (this *Document_) EndReview()  {
	retVal := this.Call(0x00000164, nil)
	_= retVal
}

func (this *Document_) StyleSheets() *StyleSheets {
	retVal := this.PropGet(0x00000168, nil)
	return NewStyleSheets(retVal.PdispValVal(), false, true)
}

func (this *Document_) DefaultTableStyle() ole.Variant {
	retVal := this.PropGet(0x0000016d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Document_) PasswordEncryptionProvider() string {
	retVal := this.PropGet(0x0000016f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) PasswordEncryptionAlgorithm() string {
	retVal := this.PropGet(0x00000170, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) PasswordEncryptionKeyLength() int32 {
	retVal := this.PropGet(0x00000171, nil)
	return retVal.LValVal()
}

func (this *Document_) PasswordEncryptionFileProperties() bool {
	retVal := this.PropGet(0x00000172, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Document__SetPasswordEncryptionOptions_OptArgs= []string{
	"PasswordEncryptionFileProperties", 
}

func (this *Document_) SetPasswordEncryptionOptions(passwordEncryptionProvider string, passwordEncryptionAlgorithm string, passwordEncryptionKeyLength int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__SetPasswordEncryptionOptions_OptArgs, optArgs)
	retVal := this.Call(0x00000169, []interface{}{passwordEncryptionProvider, passwordEncryptionAlgorithm, passwordEncryptionKeyLength}, optArgs...)
	_= retVal
}

func (this *Document_) RecheckSmartTags()  {
	retVal := this.Call(0x0000016b, nil)
	_= retVal
}

func (this *Document_) RemoveSmartTags()  {
	retVal := this.Call(0x0000016c, nil)
	_= retVal
}

func (this *Document_) SetDefaultTableStyle(style *ole.Variant, setInTemplate bool)  {
	retVal := this.Call(0x0000016e, []interface{}{style, setInTemplate})
	_= retVal
}

func (this *Document_) DeleteAllComments()  {
	retVal := this.Call(0x00000173, nil)
	_= retVal
}

func (this *Document_) AcceptAllRevisionsShown()  {
	retVal := this.Call(0x00000174, nil)
	_= retVal
}

func (this *Document_) RejectAllRevisionsShown()  {
	retVal := this.Call(0x00000175, nil)
	_= retVal
}

func (this *Document_) DeleteAllCommentsShown()  {
	retVal := this.Call(0x00000176, nil)
	_= retVal
}

func (this *Document_) ResetFormFields()  {
	retVal := this.Call(0x00000177, nil)
	_= retVal
}

var Document__SaveAs_OptArgs= []string{
	"FileName", "FileFormat", "LockComments", "Password", 
	"AddToRecentFiles", "WritePassword", "ReadOnlyRecommended", "EmbedTrueTypeFonts", 
	"SaveNativePictureFormat", "SaveFormsData", "SaveAsAOCELetter", "Encoding", 
	"InsertLineBreaks", "AllowSubstitutions", "LineEnding", "AddBiDiMarks", 
}

func (this *Document_) SaveAs(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__SaveAs_OptArgs, optArgs)
	retVal := this.Call(0x00000178, nil, optArgs...)
	_= retVal
}

func (this *Document_) EmbedLinguisticData() bool {
	retVal := this.PropGet(0x00000179, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetEmbedLinguisticData(rhs bool)  {
	retVal := this.PropPut(0x00000179, []interface{}{rhs})
	_= retVal
}

func (this *Document_) FormattingShowFont() bool {
	retVal := this.PropGet(0x000001c0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetFormattingShowFont(rhs bool)  {
	retVal := this.PropPut(0x000001c0, []interface{}{rhs})
	_= retVal
}

func (this *Document_) FormattingShowClear() bool {
	retVal := this.PropGet(0x000001c1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetFormattingShowClear(rhs bool)  {
	retVal := this.PropPut(0x000001c1, []interface{}{rhs})
	_= retVal
}

func (this *Document_) FormattingShowParagraph() bool {
	retVal := this.PropGet(0x000001c2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetFormattingShowParagraph(rhs bool)  {
	retVal := this.PropPut(0x000001c2, []interface{}{rhs})
	_= retVal
}

func (this *Document_) FormattingShowNumbering() bool {
	retVal := this.PropGet(0x000001c3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetFormattingShowNumbering(rhs bool)  {
	retVal := this.PropPut(0x000001c3, []interface{}{rhs})
	_= retVal
}

func (this *Document_) FormattingShowFilter() int32 {
	retVal := this.PropGet(0x000001c4, nil)
	return retVal.LValVal()
}

func (this *Document_) SetFormattingShowFilter(rhs int32)  {
	retVal := this.PropPut(0x000001c4, []interface{}{rhs})
	_= retVal
}

func (this *Document_) CheckNewSmartTags()  {
	retVal := this.Call(0x0000017a, nil)
	_= retVal
}

func (this *Document_) Permission() *ole.DispatchClass {
	retVal := this.PropGet(0x000001c5, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) XMLNodes() *XMLNodes {
	retVal := this.PropGet(0x000001cc, nil)
	return NewXMLNodes(retVal.PdispValVal(), false, true)
}

func (this *Document_) XMLSchemaReferences() *XMLSchemaReferences {
	retVal := this.PropGet(0x000001cd, nil)
	return NewXMLSchemaReferences(retVal.PdispValVal(), false, true)
}

func (this *Document_) SmartDocument() *ole.DispatchClass {
	retVal := this.PropGet(0x000001ce, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) SharedWorkspace() *ole.DispatchClass {
	retVal := this.PropGet(0x000001cf, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) Sync() *ole.DispatchClass {
	retVal := this.PropGet(0x000001d2, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) EnforceStyle() bool {
	retVal := this.PropGet(0x000001d7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetEnforceStyle(rhs bool)  {
	retVal := this.PropPut(0x000001d7, []interface{}{rhs})
	_= retVal
}

func (this *Document_) AutoFormatOverride() bool {
	retVal := this.PropGet(0x000001d8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetAutoFormatOverride(rhs bool)  {
	retVal := this.PropPut(0x000001d8, []interface{}{rhs})
	_= retVal
}

func (this *Document_) XMLSaveDataOnly() bool {
	retVal := this.PropGet(0x000001d9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetXMLSaveDataOnly(rhs bool)  {
	retVal := this.PropPut(0x000001d9, []interface{}{rhs})
	_= retVal
}

func (this *Document_) XMLHideNamespaces() bool {
	retVal := this.PropGet(0x000001dd, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetXMLHideNamespaces(rhs bool)  {
	retVal := this.PropPut(0x000001dd, []interface{}{rhs})
	_= retVal
}

func (this *Document_) XMLShowAdvancedErrors() bool {
	retVal := this.PropGet(0x000001de, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetXMLShowAdvancedErrors(rhs bool)  {
	retVal := this.PropPut(0x000001de, []interface{}{rhs})
	_= retVal
}

func (this *Document_) XMLUseXSLTWhenSaving() bool {
	retVal := this.PropGet(0x000001da, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetXMLUseXSLTWhenSaving(rhs bool)  {
	retVal := this.PropPut(0x000001da, []interface{}{rhs})
	_= retVal
}

func (this *Document_) XMLSaveThroughXSLT() string {
	retVal := this.PropGet(0x000001db, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) SetXMLSaveThroughXSLT(rhs string)  {
	retVal := this.PropPut(0x000001db, []interface{}{rhs})
	_= retVal
}

func (this *Document_) DocumentLibraryVersions() *ole.DispatchClass {
	retVal := this.PropGet(0x000001dc, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) ReadingModeLayoutFrozen() bool {
	retVal := this.PropGet(0x000001e1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetReadingModeLayoutFrozen(rhs bool)  {
	retVal := this.PropPut(0x000001e1, []interface{}{rhs})
	_= retVal
}

func (this *Document_) RemoveDateAndTime() bool {
	retVal := this.PropGet(0x000001e4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetRemoveDateAndTime(rhs bool)  {
	retVal := this.PropPut(0x000001e4, []interface{}{rhs})
	_= retVal
}

var Document__SendFaxOverInternet_OptArgs= []string{
	"Recipients", "Subject", "ShowMessage", 
}

func (this *Document_) SendFaxOverInternet(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__SendFaxOverInternet_OptArgs, optArgs)
	retVal := this.Call(0x000001d0, nil, optArgs...)
	_= retVal
}

func (this *Document_) TransformDocument(path string, dataOnly bool)  {
	retVal := this.Call(0x000001f4, []interface{}{path, dataOnly})
	_= retVal
}

var Document__Protect_OptArgs= []string{
	"NoReset", "Password", "UseIRM", "EnforceStyleLock", 
}

func (this *Document_) Protect(type_ int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__Protect_OptArgs, optArgs)
	retVal := this.Call(0x000001d3, []interface{}{type_}, optArgs...)
	_= retVal
}

var Document__SelectAllEditableRanges_OptArgs= []string{
	"EditorID", 
}

func (this *Document_) SelectAllEditableRanges(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__SelectAllEditableRanges_OptArgs, optArgs)
	retVal := this.Call(0x000001d4, nil, optArgs...)
	_= retVal
}

var Document__DeleteAllEditableRanges_OptArgs= []string{
	"EditorID", 
}

func (this *Document_) DeleteAllEditableRanges(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__DeleteAllEditableRanges_OptArgs, optArgs)
	retVal := this.Call(0x000001d5, nil, optArgs...)
	_= retVal
}

func (this *Document_) DeleteAllInkAnnotations()  {
	retVal := this.Call(0x000001df, nil)
	_= retVal
}

func (this *Document_) AddDocumentWorkspaceHeader(richFormat bool, url string, title string, description string, id string)  {
	retVal := this.Call(0x000001e2, []interface{}{richFormat, url, title, description, id})
	_= retVal
}

func (this *Document_) RemoveDocumentWorkspaceHeader(id string)  {
	retVal := this.Call(0x000001e3, []interface{}{id})
	_= retVal
}

var Document__Compare_OptArgs= []string{
	"AuthorName", "CompareTarget", "DetectFormatChanges", "IgnoreAllComparisonWarnings", 
	"AddToRecentFiles", "RemovePersonalInformation", "RemoveDateAndTime", 
}

func (this *Document_) Compare(name string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__Compare_OptArgs, optArgs)
	retVal := this.Call(0x000001e5, []interface{}{name}, optArgs...)
	_= retVal
}

func (this *Document_) RemoveLockedStyles()  {
	retVal := this.Call(0x000001e7, nil)
	_= retVal
}

func (this *Document_) ChildNodeSuggestions() *XMLChildNodeSuggestions {
	retVal := this.PropGet(0x000001e6, nil)
	return NewXMLChildNodeSuggestions(retVal.PdispValVal(), false, true)
}

func (this *Document_) SelectSingleNode(xpath string, prefixMapping string, fastSearchSkippingTextNodes bool) *XMLNode {
	retVal := this.Call(0x000001e8, []interface{}{xpath, prefixMapping, fastSearchSkippingTextNodes})
	return NewXMLNode(retVal.PdispValVal(), false, true)
}

func (this *Document_) SelectNodes(xpath string, prefixMapping string, fastSearchSkippingTextNodes bool) *XMLNodes {
	retVal := this.Call(0x000001e9, []interface{}{xpath, prefixMapping, fastSearchSkippingTextNodes})
	return NewXMLNodes(retVal.PdispValVal(), false, true)
}

func (this *Document_) XMLSchemaViolations() *XMLNodes {
	retVal := this.PropGet(0x000001ea, nil)
	return NewXMLNodes(retVal.PdispValVal(), false, true)
}

func (this *Document_) ReadingLayoutSizeX() int32 {
	retVal := this.PropGet(0x000001eb, nil)
	return retVal.LValVal()
}

func (this *Document_) SetReadingLayoutSizeX(rhs int32)  {
	retVal := this.PropPut(0x000001eb, []interface{}{rhs})
	_= retVal
}

func (this *Document_) ReadingLayoutSizeY() int32 {
	retVal := this.PropGet(0x000001ec, nil)
	return retVal.LValVal()
}

func (this *Document_) SetReadingLayoutSizeY(rhs int32)  {
	retVal := this.PropPut(0x000001ec, []interface{}{rhs})
	_= retVal
}

func (this *Document_) StyleSortMethod() int32 {
	retVal := this.PropGet(0x000001ed, nil)
	return retVal.LValVal()
}

func (this *Document_) SetStyleSortMethod(rhs int32)  {
	retVal := this.PropPut(0x000001ed, []interface{}{rhs})
	_= retVal
}

func (this *Document_) ContentTypeProperties() *ole.DispatchClass {
	retVal := this.PropGet(0x000001f0, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) TrackMoves() bool {
	retVal := this.PropGet(0x000001f3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetTrackMoves(rhs bool)  {
	retVal := this.PropPut(0x000001f3, []interface{}{rhs})
	_= retVal
}

func (this *Document_) TrackFormatting() bool {
	retVal := this.PropGet(0x000001f6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetTrackFormatting(rhs bool)  {
	retVal := this.PropPut(0x000001f6, []interface{}{rhs})
	_= retVal
}

func (this *Document_) Dummy1()  {
	retVal := this.PropGet(0x000001f7, nil)
	_= retVal
}

func (this *Document_) OMaths() *OMaths {
	retVal := this.PropGet(0x000001f8, nil)
	return NewOMaths(retVal.PdispValVal(), false, true)
}

func (this *Document_) RemoveDocumentInformation(removeDocInfoType int32)  {
	retVal := this.Call(0x000001ef, []interface{}{removeDocInfoType})
	_= retVal
}

var Document__CheckInWithVersion_OptArgs= []string{
	"VersionType", 
}

func (this *Document_) CheckInWithVersion(saveChanges bool, comments *ole.Variant, makePublic bool, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__CheckInWithVersion_OptArgs, optArgs)
	retVal := this.Call(0x000001f5, []interface{}{saveChanges, comments, makePublic}, optArgs...)
	_= retVal
}

func (this *Document_) Dummy2()  {
	retVal := this.Call(0x000001f9, nil)
	_= retVal
}

func (this *Document_) Dummy3()  {
	retVal := this.PropGet(0x000001fa, nil)
	_= retVal
}

func (this *Document_) ServerPolicy() *ole.DispatchClass {
	retVal := this.PropGet(0x000001fb, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) ContentControls() *ContentControls {
	retVal := this.PropGet(0x000001fc, nil)
	return NewContentControls(retVal.PdispValVal(), false, true)
}

func (this *Document_) DocumentInspectors() *ole.DispatchClass {
	retVal := this.PropGet(0x000001fe, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) LockServerFile()  {
	retVal := this.Call(0x000001fd, nil)
	_= retVal
}

func (this *Document_) GetWorkflowTasks() *ole.DispatchClass {
	retVal := this.Call(0x000001ff, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) GetWorkflowTemplates() *ole.DispatchClass {
	retVal := this.Call(0x00000200, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) Dummy4()  {
	retVal := this.Call(0x00000202, nil)
	_= retVal
}

func (this *Document_) AddMeetingWorkspaceHeader(skipIfAbsent bool, url string, title string, description string, id string)  {
	retVal := this.Call(0x00000203, []interface{}{skipIfAbsent, url, title, description, id})
	_= retVal
}

func (this *Document_) Bibliography() *Bibliography {
	retVal := this.PropGet(0x00000204, nil)
	return NewBibliography(retVal.PdispValVal(), false, true)
}

func (this *Document_) LockTheme() bool {
	retVal := this.PropGet(0x00000205, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetLockTheme(rhs bool)  {
	retVal := this.PropPut(0x00000205, []interface{}{rhs})
	_= retVal
}

func (this *Document_) LockQuickStyleSet() bool {
	retVal := this.PropGet(0x00000206, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetLockQuickStyleSet(rhs bool)  {
	retVal := this.PropPut(0x00000206, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OriginalDocumentTitle() string {
	retVal := this.PropGet(0x00000207, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) RevisedDocumentTitle() string {
	retVal := this.PropGet(0x00000208, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) CustomXMLParts() *ole.DispatchClass {
	retVal := this.PropGet(0x00000209, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) FormattingShowNextLevel() bool {
	retVal := this.PropGet(0x0000020a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetFormattingShowNextLevel(rhs bool)  {
	retVal := this.PropPut(0x0000020a, []interface{}{rhs})
	_= retVal
}

func (this *Document_) FormattingShowUserStyleName() bool {
	retVal := this.PropGet(0x0000020b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetFormattingShowUserStyleName(rhs bool)  {
	retVal := this.PropPut(0x0000020b, []interface{}{rhs})
	_= retVal
}

func (this *Document_) SaveAsQuickStyleSet(fileName string)  {
	retVal := this.Call(0x0000020c, []interface{}{fileName})
	_= retVal
}

func (this *Document_) ApplyQuickStyleSet(name string)  {
	retVal := this.Call(0x0000020d, []interface{}{name})
	_= retVal
}

func (this *Document_) Research() *Research {
	retVal := this.PropGet(0x0000020e, nil)
	return NewResearch(retVal.PdispValVal(), false, true)
}

func (this *Document_) Final() bool {
	retVal := this.PropGet(0x0000020f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetFinal(rhs bool)  {
	retVal := this.PropPut(0x0000020f, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OMathBreakBin() int32 {
	retVal := this.PropGet(0x00000210, nil)
	return retVal.LValVal()
}

func (this *Document_) SetOMathBreakBin(rhs int32)  {
	retVal := this.PropPut(0x00000210, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OMathBreakSub() int32 {
	retVal := this.PropGet(0x00000211, nil)
	return retVal.LValVal()
}

func (this *Document_) SetOMathBreakSub(rhs int32)  {
	retVal := this.PropPut(0x00000211, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OMathJc() int32 {
	retVal := this.PropGet(0x00000212, nil)
	return retVal.LValVal()
}

func (this *Document_) SetOMathJc(rhs int32)  {
	retVal := this.PropPut(0x00000212, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OMathLeftMargin() float32 {
	retVal := this.PropGet(0x00000213, nil)
	return retVal.FltValVal()
}

func (this *Document_) SetOMathLeftMargin(rhs float32)  {
	retVal := this.PropPut(0x00000213, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OMathRightMargin() float32 {
	retVal := this.PropGet(0x00000214, nil)
	return retVal.FltValVal()
}

func (this *Document_) SetOMathRightMargin(rhs float32)  {
	retVal := this.PropPut(0x00000214, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OMathWrap() float32 {
	retVal := this.PropGet(0x00000217, nil)
	return retVal.FltValVal()
}

func (this *Document_) SetOMathWrap(rhs float32)  {
	retVal := this.PropPut(0x00000217, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OMathIntSubSupLim() bool {
	retVal := this.PropGet(0x00000218, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetOMathIntSubSupLim(rhs bool)  {
	retVal := this.PropPut(0x00000218, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OMathNarySupSubLim() bool {
	retVal := this.PropGet(0x00000219, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetOMathNarySupSubLim(rhs bool)  {
	retVal := this.PropPut(0x00000219, []interface{}{rhs})
	_= retVal
}

func (this *Document_) OMathSmallFrac() bool {
	retVal := this.PropGet(0x0000021b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetOMathSmallFrac(rhs bool)  {
	retVal := this.PropPut(0x0000021b, []interface{}{rhs})
	_= retVal
}

func (this *Document_) WordOpenXML() string {
	retVal := this.PropGet(0x0000021e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) DocumentTheme() *ole.DispatchClass {
	retVal := this.PropGet(0x00000221, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Document_) ApplyDocumentTheme(fileName string)  {
	retVal := this.Call(0x00000222, []interface{}{fileName})
	_= retVal
}

func (this *Document_) HasVBProject() bool {
	retVal := this.PropGet(0x00000224, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SelectLinkedControls(node *ole.DispatchClass) *ContentControls {
	retVal := this.Call(0x00000225, []interface{}{node})
	return NewContentControls(retVal.PdispValVal(), false, true)
}

func (this *Document_) SelectUnlinkedControls(stream *ole.DispatchClass) *ContentControls {
	retVal := this.Call(0x00000226, []interface{}{stream})
	return NewContentControls(retVal.PdispValVal(), false, true)
}

func (this *Document_) SelectContentControlsByTitle(title string) *ContentControls {
	retVal := this.Call(0x00000227, []interface{}{title})
	return NewContentControls(retVal.PdispValVal(), false, true)
}

var Document__ExportAsFixedFormat_OptArgs= []string{
	"FixedFormatExtClassPtr", 
}

func (this *Document_) ExportAsFixedFormat(outputFileName string, exportFormat int32, openAfterExport bool, optimizeFor int32, range_ int32, from int32, to int32, item int32, includeDocProps bool, keepIRM bool, createBookmarks int32, docStructureTags bool, bitmapMissingFonts bool, useISO19005_1 bool, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__ExportAsFixedFormat_OptArgs, optArgs)
	retVal := this.Call(0x00000228, []interface{}{outputFileName, exportFormat, openAfterExport, optimizeFor, range_, from, to, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1}, optArgs...)
	_= retVal
}

func (this *Document_) FreezeLayout()  {
	retVal := this.Call(0x00000229, nil)
	_= retVal
}

func (this *Document_) UnfreezeLayout()  {
	retVal := this.Call(0x0000022a, nil)
	_= retVal
}

func (this *Document_) OMathFontName() string {
	retVal := this.PropGet(0x0000022b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) SetOMathFontName(rhs string)  {
	retVal := this.PropPut(0x0000022b, []interface{}{rhs})
	_= retVal
}

func (this *Document_) DowngradeDocument()  {
	retVal := this.Call(0x0000022e, nil)
	_= retVal
}

func (this *Document_) EncryptionProvider() string {
	retVal := this.PropGet(0x0000022f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Document_) SetEncryptionProvider(rhs string)  {
	retVal := this.PropPut(0x0000022f, []interface{}{rhs})
	_= retVal
}

func (this *Document_) UseMathDefaults() bool {
	retVal := this.PropGet(0x00000230, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Document_) SetUseMathDefaults(rhs bool)  {
	retVal := this.PropPut(0x00000230, []interface{}{rhs})
	_= retVal
}

func (this *Document_) CurrentRsid() int32 {
	retVal := this.PropGet(0x00000233, nil)
	return retVal.LValVal()
}

func (this *Document_) Convert()  {
	retVal := this.Call(0x00000231, nil)
	_= retVal
}

func (this *Document_) SelectContentControlsByTag(tag string) *ContentControls {
	retVal := this.Call(0x00000232, []interface{}{tag})
	return NewContentControls(retVal.PdispValVal(), false, true)
}

func (this *Document_) ConvertAutoHyphens()  {
	retVal := this.Call(0x0000028a, nil)
	_= retVal
}

func (this *Document_) DocID() int32 {
	retVal := this.PropGet(0x00000234, nil)
	return retVal.LValVal()
}

func (this *Document_) ApplyQuickStyleSet2(style *ole.Variant)  {
	retVal := this.Call(0x00000236, []interface{}{style})
	_= retVal
}

func (this *Document_) CompatibilityMode() int32 {
	retVal := this.PropGet(0x00000237, nil)
	return retVal.LValVal()
}

var Document__SaveAs2_OptArgs= []string{
	"FileName", "FileFormat", "LockComments", "Password", 
	"AddToRecentFiles", "WritePassword", "ReadOnlyRecommended", "EmbedTrueTypeFonts", 
	"SaveNativePictureFormat", "SaveFormsData", "SaveAsAOCELetter", "Encoding", 
	"InsertLineBreaks", "AllowSubstitutions", "LineEnding", "AddBiDiMarks", "CompatibilityMode", 
}

func (this *Document_) SaveAs2(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Document__SaveAs2_OptArgs, optArgs)
	retVal := this.Call(0x00000238, nil, optArgs...)
	_= retVal
}

func (this *Document_) CoAuthoring() *CoAuthoring {
	retVal := this.PropGet(0x00000258, nil)
	return NewCoAuthoring(retVal.PdispValVal(), false, true)
}

func (this *Document_) SetCompatibilityMode(mode int32)  {
	retVal := this.Call(0x0000023b, []interface{}{mode})
	_= retVal
}

