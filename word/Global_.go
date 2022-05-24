package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209B9-0000-0000-C000-000000000046
var IID_Global_ = syscall.GUID{0x000209B9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Global_ struct {
	ole.OleClient
}

func NewGlobal_(pDisp *win32.IDispatch, addRef bool, scoped bool) *Global_ {
	 if pDisp == nil {
		return nil;
	}
	p := &Global_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Global_FromVar(v ole.Variant) *Global_ {
	return NewGlobal_(v.IDispatch(), false, false)
}

func (this *Global_) IID() *syscall.GUID {
	return &IID_Global_
}

func (this *Global_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Global_) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Global_) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Global_) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Global_) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Global_) Documents() *Documents {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewDocuments(retVal.IDispatch(), false, true)
}

func (this *Global_) Windows() *Windows {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewWindows(retVal.IDispatch(), false, true)
}

func (this *Global_) ActiveDocument() *Document {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewDocument(retVal.IDispatch(), false, true)
}

func (this *Global_) ActiveWindow() *Window {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Global_) Selection() *Selection {
	retVal, _ := this.PropGet(0x00000005, nil)
	return NewSelection(retVal.IDispatch(), false, true)
}

func (this *Global_) WordBasic() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000006, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Global_) PrintPreview() bool {
	retVal, _ := this.PropGet(0x0000001b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Global_) SetPrintPreview(rhs bool)  {
	_ = this.PropPut(0x0000001b, []interface{}{rhs})
}

func (this *Global_) RecentFiles() *RecentFiles {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewRecentFiles(retVal.IDispatch(), false, true)
}

func (this *Global_) NormalTemplate() *Template {
	retVal, _ := this.PropGet(0x00000008, nil)
	return NewTemplate(retVal.IDispatch(), false, true)
}

func (this *Global_) System() *System {
	retVal, _ := this.PropGet(0x00000009, nil)
	return NewSystem(retVal.IDispatch(), false, true)
}

func (this *Global_) AutoCorrect() *AutoCorrect {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewAutoCorrect(retVal.IDispatch(), false, true)
}

func (this *Global_) FontNames() *FontNames {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return NewFontNames(retVal.IDispatch(), false, true)
}

func (this *Global_) LandscapeFontNames() *FontNames {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return NewFontNames(retVal.IDispatch(), false, true)
}

func (this *Global_) PortraitFontNames() *FontNames {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return NewFontNames(retVal.IDispatch(), false, true)
}

func (this *Global_) Languages() *Languages {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return NewLanguages(retVal.IDispatch(), false, true)
}

func (this *Global_) Assistant() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Global_) FileConverters() *FileConverters {
	retVal, _ := this.PropGet(0x00000011, nil)
	return NewFileConverters(retVal.IDispatch(), false, true)
}

func (this *Global_) Dialogs() *Dialogs {
	retVal, _ := this.PropGet(0x00000013, nil)
	return NewDialogs(retVal.IDispatch(), false, true)
}

func (this *Global_) CaptionLabels() *CaptionLabels {
	retVal, _ := this.PropGet(0x00000014, nil)
	return NewCaptionLabels(retVal.IDispatch(), false, true)
}

func (this *Global_) AutoCaptions() *AutoCaptions {
	retVal, _ := this.PropGet(0x00000015, nil)
	return NewAutoCaptions(retVal.IDispatch(), false, true)
}

func (this *Global_) AddIns() *AddIns {
	retVal, _ := this.PropGet(0x00000016, nil)
	return NewAddIns(retVal.IDispatch(), false, true)
}

func (this *Global_) Tasks() *Tasks {
	retVal, _ := this.PropGet(0x0000001c, nil)
	return NewTasks(retVal.IDispatch(), false, true)
}

func (this *Global_) MacroContainer() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000037, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Global_) CommandBars() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000039, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Global__SynonymInfo_OptArgs= []string{
	"LanguageID", 
}

func (this *Global_) SynonymInfo(word string, optArgs ...interface{}) *SynonymInfo {
	optArgs = ole.ProcessOptArgs(Global__SynonymInfo_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000003b, []interface{}{word}, optArgs...)
	return NewSynonymInfo(retVal.IDispatch(), false, true)
}

func (this *Global_) VBE() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000003d, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Global_) ListGalleries() *ListGalleries {
	retVal, _ := this.PropGet(0x00000041, nil)
	return NewListGalleries(retVal.IDispatch(), false, true)
}

func (this *Global_) ActivePrinter() string {
	retVal, _ := this.PropGet(0x00000042, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Global_) SetActivePrinter(rhs string)  {
	_ = this.PropPut(0x00000042, []interface{}{rhs})
}

func (this *Global_) Templates() *Templates {
	retVal, _ := this.PropGet(0x00000043, nil)
	return NewTemplates(retVal.IDispatch(), false, true)
}

func (this *Global_) CustomizationContext() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000044, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Global_) SetCustomizationContext(rhs *win32.IUnknown)  {
	_ = this.PropPut(0x00000044, []interface{}{rhs})
}

func (this *Global_) KeyBindings() *KeyBindings {
	retVal, _ := this.PropGet(0x00000045, nil)
	return NewKeyBindings(retVal.IDispatch(), false, true)
}

var Global__KeysBoundTo_OptArgs= []string{
	"CommandParameter", 
}

func (this *Global_) KeysBoundTo(keyCategory int32, command string, optArgs ...interface{}) *KeysBoundTo {
	optArgs = ole.ProcessOptArgs(Global__KeysBoundTo_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000046, []interface{}{keyCategory, command}, optArgs...)
	return NewKeysBoundTo(retVal.IDispatch(), false, true)
}

var Global__FindKey_OptArgs= []string{
	"KeyCode2", 
}

func (this *Global_) FindKey(keyCode int32, optArgs ...interface{}) *KeyBinding {
	optArgs = ole.ProcessOptArgs(Global__FindKey_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000047, []interface{}{keyCode}, optArgs...)
	return NewKeyBinding(retVal.IDispatch(), false, true)
}

func (this *Global_) Options() *Options {
	retVal, _ := this.PropGet(0x0000005d, nil)
	return NewOptions(retVal.IDispatch(), false, true)
}

func (this *Global_) CustomDictionaries() *Dictionaries {
	retVal, _ := this.PropGet(0x0000005f, nil)
	return NewDictionaries(retVal.IDispatch(), false, true)
}

func (this *Global_) SetStatusBar(rhs string)  {
	_ = this.PropPut(0x00000061, []interface{}{rhs})
}

func (this *Global_) ShowVisualBasicEditor() bool {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Global_) SetShowVisualBasicEditor(rhs bool)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *Global_) IsObjectValid(object *win32.IUnknown) bool {
	retVal, _ := this.PropGet(0x0000006d, []interface{}{object})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Global_) HangulHanjaDictionaries() *HangulHanjaConversionDictionaries {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return NewHangulHanjaConversionDictionaries(retVal.IDispatch(), false, true)
}

var Global__Repeat_OptArgs= []string{
	"Times", 
}

func (this *Global_) Repeat(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Global__Repeat_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000131, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Global_) DDEExecute(channel int32, command string)  {
	retVal, _ := this.Call(0x00000136, []interface{}{channel, command})
	_= retVal
}

func (this *Global_) DDEInitiate(app string, topic string) int32 {
	retVal, _ := this.Call(0x00000137, []interface{}{app, topic})
	return retVal.LValVal()
}

func (this *Global_) DDEPoke(channel int32, item string, data string)  {
	retVal, _ := this.Call(0x00000138, []interface{}{channel, item, data})
	_= retVal
}

func (this *Global_) DDERequest(channel int32, item string) string {
	retVal, _ := this.Call(0x00000139, []interface{}{channel, item})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Global_) DDETerminate(channel int32)  {
	retVal, _ := this.Call(0x0000013a, []interface{}{channel})
	_= retVal
}

func (this *Global_) DDETerminateAll()  {
	retVal, _ := this.Call(0x0000013b, nil)
	_= retVal
}

var Global__BuildKeyCode_OptArgs= []string{
	"Arg2", "Arg3", "Arg4", 
}

func (this *Global_) BuildKeyCode(arg1 int32, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Global__BuildKeyCode_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000013c, []interface{}{arg1}, optArgs...)
	return retVal.LValVal()
}

var Global__KeyString_OptArgs= []string{
	"KeyCode2", 
}

func (this *Global_) KeyString(keyCode int32, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Global__KeyString_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000013d, []interface{}{keyCode}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var Global__CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "MainDictionary", "CustomDictionary2", 
	"CustomDictionary3", "CustomDictionary4", "CustomDictionary5", "CustomDictionary6", 
	"CustomDictionary7", "CustomDictionary8", "CustomDictionary9", "CustomDictionary10", 
}

func (this *Global_) CheckSpelling(word string, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Global__CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000144, []interface{}{word}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Global__GetSpellingSuggestions_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "MainDictionary", "SuggestionMode", 
	"CustomDictionary2", "CustomDictionary3", "CustomDictionary4", "CustomDictionary5", 
	"CustomDictionary6", "CustomDictionary7", "CustomDictionary8", "CustomDictionary9", "CustomDictionary10", 
}

func (this *Global_) GetSpellingSuggestions(word string, optArgs ...interface{}) *SpellingSuggestions {
	optArgs = ole.ProcessOptArgs(Global__GetSpellingSuggestions_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000147, []interface{}{word}, optArgs...)
	return NewSpellingSuggestions(retVal.IDispatch(), false, true)
}

func (this *Global_) Help(helpType *ole.Variant)  {
	retVal, _ := this.Call(0x00000149, []interface{}{helpType})
	_= retVal
}

func (this *Global_) NewWindow() *Window {
	retVal, _ := this.Call(0x00000159, nil)
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Global_) CleanString(string string) string {
	retVal, _ := this.Call(0x00000162, []interface{}{string})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Global_) ChangeFileOpenDirectory(path string)  {
	retVal, _ := this.Call(0x00000163, []interface{}{path})
	_= retVal
}

func (this *Global_) InchesToPoints(inches float32) float32 {
	retVal, _ := this.Call(0x00000172, []interface{}{inches})
	return retVal.FltValVal()
}

func (this *Global_) CentimetersToPoints(centimeters float32) float32 {
	retVal, _ := this.Call(0x00000173, []interface{}{centimeters})
	return retVal.FltValVal()
}

func (this *Global_) MillimetersToPoints(millimeters float32) float32 {
	retVal, _ := this.Call(0x00000174, []interface{}{millimeters})
	return retVal.FltValVal()
}

func (this *Global_) PicasToPoints(picas float32) float32 {
	retVal, _ := this.Call(0x00000175, []interface{}{picas})
	return retVal.FltValVal()
}

func (this *Global_) LinesToPoints(lines float32) float32 {
	retVal, _ := this.Call(0x00000176, []interface{}{lines})
	return retVal.FltValVal()
}

func (this *Global_) PointsToInches(points float32) float32 {
	retVal, _ := this.Call(0x0000017c, []interface{}{points})
	return retVal.FltValVal()
}

func (this *Global_) PointsToCentimeters(points float32) float32 {
	retVal, _ := this.Call(0x0000017d, []interface{}{points})
	return retVal.FltValVal()
}

func (this *Global_) PointsToMillimeters(points float32) float32 {
	retVal, _ := this.Call(0x0000017e, []interface{}{points})
	return retVal.FltValVal()
}

func (this *Global_) PointsToPicas(points float32) float32 {
	retVal, _ := this.Call(0x0000017f, []interface{}{points})
	return retVal.FltValVal()
}

func (this *Global_) PointsToLines(points float32) float32 {
	retVal, _ := this.Call(0x00000180, []interface{}{points})
	return retVal.FltValVal()
}

var Global__PointsToPixels_OptArgs= []string{
	"fVertical", 
}

func (this *Global_) PointsToPixels(points float32, optArgs ...interface{}) float32 {
	optArgs = ole.ProcessOptArgs(Global__PointsToPixels_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000181, []interface{}{points}, optArgs...)
	return retVal.FltValVal()
}

var Global__PixelsToPoints_OptArgs= []string{
	"fVertical", 
}

func (this *Global_) PixelsToPoints(pixels float32, optArgs ...interface{}) float32 {
	optArgs = ole.ProcessOptArgs(Global__PixelsToPoints_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000182, []interface{}{pixels}, optArgs...)
	return retVal.FltValVal()
}

func (this *Global_) LanguageSettings() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Global_) AnswerWizard() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000070, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Global_) AutoCorrectEmail() *AutoCorrect {
	retVal, _ := this.PropGet(0x00000071, nil)
	return NewAutoCorrect(retVal.IDispatch(), false, true)
}

func (this *Global_) ProtectedViewWindows() *ProtectedViewWindows {
	retVal, _ := this.PropGet(0x00000072, nil)
	return NewProtectedViewWindows(retVal.IDispatch(), false, true)
}

func (this *Global_) ActiveProtectedViewWindow() *ProtectedViewWindow {
	retVal, _ := this.PropGet(0x00000073, nil)
	return NewProtectedViewWindow(retVal.IDispatch(), false, true)
}

func (this *Global_) IsSandboxed() bool {
	retVal, _ := this.PropGet(0x00000074, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

