package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020970-0000-0000-C000-000000000046
var IID_Application_ = syscall.GUID{0x00020970, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Application_ struct {
	ole.OleClient
}

func NewApplication_(pDisp *win32.IDispatch, addRef bool, scoped bool) *Application_ {
	 if pDisp == nil {
		return nil;
	}
	p := &Application_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Application_FromVar(v ole.Variant) *Application_ {
	return NewApplication_(v.IDispatch(), false, false)
}

func (this *Application_) IID() *syscall.GUID {
	return &IID_Application_
}

func (this *Application_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Application_) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Application_) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Application_) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) Documents() *Documents {
	retVal, _ := this.PropGet(0x00000006, nil)
	return NewDocuments(retVal.IDispatch(), false, true)
}

func (this *Application_) Windows() *Windows {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewWindows(retVal.IDispatch(), false, true)
}

func (this *Application_) ActiveDocument() *Document {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewDocument(retVal.IDispatch(), false, true)
}

func (this *Application_) ActiveWindow() *Window {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Application_) Selection() *Selection {
	retVal, _ := this.PropGet(0x00000005, nil)
	return NewSelection(retVal.IDispatch(), false, true)
}

func (this *Application_) WordBasic() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) RecentFiles() *RecentFiles {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewRecentFiles(retVal.IDispatch(), false, true)
}

func (this *Application_) NormalTemplate() *Template {
	retVal, _ := this.PropGet(0x00000008, nil)
	return NewTemplate(retVal.IDispatch(), false, true)
}

func (this *Application_) System() *System {
	retVal, _ := this.PropGet(0x00000009, nil)
	return NewSystem(retVal.IDispatch(), false, true)
}

func (this *Application_) AutoCorrect() *AutoCorrect {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewAutoCorrect(retVal.IDispatch(), false, true)
}

func (this *Application_) FontNames() *FontNames {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return NewFontNames(retVal.IDispatch(), false, true)
}

func (this *Application_) LandscapeFontNames() *FontNames {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return NewFontNames(retVal.IDispatch(), false, true)
}

func (this *Application_) PortraitFontNames() *FontNames {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return NewFontNames(retVal.IDispatch(), false, true)
}

func (this *Application_) Languages() *Languages {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return NewLanguages(retVal.IDispatch(), false, true)
}

func (this *Application_) Assistant() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) Browser() *Browser {
	retVal, _ := this.PropGet(0x00000010, nil)
	return NewBrowser(retVal.IDispatch(), false, true)
}

func (this *Application_) FileConverters() *FileConverters {
	retVal, _ := this.PropGet(0x00000011, nil)
	return NewFileConverters(retVal.IDispatch(), false, true)
}

func (this *Application_) MailingLabel() *MailingLabel {
	retVal, _ := this.PropGet(0x00000012, nil)
	return NewMailingLabel(retVal.IDispatch(), false, true)
}

func (this *Application_) Dialogs() *Dialogs {
	retVal, _ := this.PropGet(0x00000013, nil)
	return NewDialogs(retVal.IDispatch(), false, true)
}

func (this *Application_) CaptionLabels() *CaptionLabels {
	retVal, _ := this.PropGet(0x00000014, nil)
	return NewCaptionLabels(retVal.IDispatch(), false, true)
}

func (this *Application_) AutoCaptions() *AutoCaptions {
	retVal, _ := this.PropGet(0x00000015, nil)
	return NewAutoCaptions(retVal.IDispatch(), false, true)
}

func (this *Application_) AddIns() *AddIns {
	retVal, _ := this.PropGet(0x00000016, nil)
	return NewAddIns(retVal.IDispatch(), false, true)
}

func (this *Application_) Visible() bool {
	retVal, _ := this.PropGet(0x00000017, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetVisible(rhs bool)  {
	_ = this.PropPut(0x00000017, []interface{}{rhs})
}

func (this *Application_) Version() string {
	retVal, _ := this.PropGet(0x00000018, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) ScreenUpdating() bool {
	retVal, _ := this.PropGet(0x0000001a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetScreenUpdating(rhs bool)  {
	_ = this.PropPut(0x0000001a, []interface{}{rhs})
}

func (this *Application_) PrintPreview() bool {
	retVal, _ := this.PropGet(0x0000001b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetPrintPreview(rhs bool)  {
	_ = this.PropPut(0x0000001b, []interface{}{rhs})
}

func (this *Application_) Tasks() *Tasks {
	retVal, _ := this.PropGet(0x0000001c, nil)
	return NewTasks(retVal.IDispatch(), false, true)
}

func (this *Application_) DisplayStatusBar() bool {
	retVal, _ := this.PropGet(0x0000001d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayStatusBar(rhs bool)  {
	_ = this.PropPut(0x0000001d, []interface{}{rhs})
}

func (this *Application_) SpecialMode() bool {
	retVal, _ := this.PropGet(0x0000001e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) UsableWidth() int32 {
	retVal, _ := this.PropGet(0x00000021, nil)
	return retVal.LValVal()
}

func (this *Application_) UsableHeight() int32 {
	retVal, _ := this.PropGet(0x00000022, nil)
	return retVal.LValVal()
}

func (this *Application_) MathCoprocessorAvailable() bool {
	retVal, _ := this.PropGet(0x00000024, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) MouseAvailable() bool {
	retVal, _ := this.PropGet(0x00000025, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) International(index int32) ole.Variant {
	retVal, _ := this.PropGet(0x0000002e, []interface{}{index})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Application_) Build() string {
	retVal, _ := this.PropGet(0x0000002f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) CapsLock() bool {
	retVal, _ := this.PropGet(0x00000030, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) NumLock() bool {
	retVal, _ := this.PropGet(0x00000031, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) UserName() string {
	retVal, _ := this.PropGet(0x00000034, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetUserName(rhs string)  {
	_ = this.PropPut(0x00000034, []interface{}{rhs})
}

func (this *Application_) UserInitials() string {
	retVal, _ := this.PropGet(0x00000035, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetUserInitials(rhs string)  {
	_ = this.PropPut(0x00000035, []interface{}{rhs})
}

func (this *Application_) UserAddress() string {
	retVal, _ := this.PropGet(0x00000036, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetUserAddress(rhs string)  {
	_ = this.PropPut(0x00000036, []interface{}{rhs})
}

func (this *Application_) MacroContainer() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000037, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) DisplayRecentFiles() bool {
	retVal, _ := this.PropGet(0x00000038, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayRecentFiles(rhs bool)  {
	_ = this.PropPut(0x00000038, []interface{}{rhs})
}

func (this *Application_) CommandBars() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000039, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Application__SynonymInfo_OptArgs= []string{
	"LanguageID", 
}

func (this *Application_) SynonymInfo(word string, optArgs ...interface{}) *SynonymInfo {
	optArgs = ole.ProcessOptArgs(Application__SynonymInfo_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x0000003b, []interface{}{word}, optArgs...)
	return NewSynonymInfo(retVal.IDispatch(), false, true)
}

func (this *Application_) VBE() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000003d, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) DefaultSaveFormat() string {
	retVal, _ := this.PropGet(0x00000040, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetDefaultSaveFormat(rhs string)  {
	_ = this.PropPut(0x00000040, []interface{}{rhs})
}

func (this *Application_) ListGalleries() *ListGalleries {
	retVal, _ := this.PropGet(0x00000041, nil)
	return NewListGalleries(retVal.IDispatch(), false, true)
}

func (this *Application_) ActivePrinter() string {
	retVal, _ := this.PropGet(0x00000042, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetActivePrinter(rhs string)  {
	_ = this.PropPut(0x00000042, []interface{}{rhs})
}

func (this *Application_) Templates() *Templates {
	retVal, _ := this.PropGet(0x00000043, nil)
	return NewTemplates(retVal.IDispatch(), false, true)
}

func (this *Application_) CustomizationContext() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000044, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) SetCustomizationContext(rhs *win32.IUnknown)  {
	_ = this.PropPut(0x00000044, []interface{}{rhs})
}

func (this *Application_) KeyBindings() *KeyBindings {
	retVal, _ := this.PropGet(0x00000045, nil)
	return NewKeyBindings(retVal.IDispatch(), false, true)
}

var Application__KeysBoundTo_OptArgs= []string{
	"CommandParameter", 
}

func (this *Application_) KeysBoundTo(keyCategory int32, command string, optArgs ...interface{}) *KeysBoundTo {
	optArgs = ole.ProcessOptArgs(Application__KeysBoundTo_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000046, []interface{}{keyCategory, command}, optArgs...)
	return NewKeysBoundTo(retVal.IDispatch(), false, true)
}

var Application__FindKey_OptArgs= []string{
	"KeyCode2", 
}

func (this *Application_) FindKey(keyCode int32, optArgs ...interface{}) *KeyBinding {
	optArgs = ole.ProcessOptArgs(Application__FindKey_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000047, []interface{}{keyCode}, optArgs...)
	return NewKeyBinding(retVal.IDispatch(), false, true)
}

func (this *Application_) Caption() string {
	retVal, _ := this.PropGet(0x00000050, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetCaption(rhs string)  {
	_ = this.PropPut(0x00000050, []interface{}{rhs})
}

func (this *Application_) Path() string {
	retVal, _ := this.PropGet(0x00000051, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) DisplayScrollBars() bool {
	retVal, _ := this.PropGet(0x00000052, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayScrollBars(rhs bool)  {
	_ = this.PropPut(0x00000052, []interface{}{rhs})
}

func (this *Application_) StartupPath() string {
	retVal, _ := this.PropGet(0x00000053, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetStartupPath(rhs string)  {
	_ = this.PropPut(0x00000053, []interface{}{rhs})
}

func (this *Application_) BackgroundSavingStatus() int32 {
	retVal, _ := this.PropGet(0x00000055, nil)
	return retVal.LValVal()
}

func (this *Application_) BackgroundPrintingStatus() int32 {
	retVal, _ := this.PropGet(0x00000056, nil)
	return retVal.LValVal()
}

func (this *Application_) Left() int32 {
	retVal, _ := this.PropGet(0x00000057, nil)
	return retVal.LValVal()
}

func (this *Application_) SetLeft(rhs int32)  {
	_ = this.PropPut(0x00000057, []interface{}{rhs})
}

func (this *Application_) Top() int32 {
	retVal, _ := this.PropGet(0x00000058, nil)
	return retVal.LValVal()
}

func (this *Application_) SetTop(rhs int32)  {
	_ = this.PropPut(0x00000058, []interface{}{rhs})
}

func (this *Application_) Width() int32 {
	retVal, _ := this.PropGet(0x00000059, nil)
	return retVal.LValVal()
}

func (this *Application_) SetWidth(rhs int32)  {
	_ = this.PropPut(0x00000059, []interface{}{rhs})
}

func (this *Application_) Height() int32 {
	retVal, _ := this.PropGet(0x0000005a, nil)
	return retVal.LValVal()
}

func (this *Application_) SetHeight(rhs int32)  {
	_ = this.PropPut(0x0000005a, []interface{}{rhs})
}

func (this *Application_) WindowState() int32 {
	retVal, _ := this.PropGet(0x0000005b, nil)
	return retVal.LValVal()
}

func (this *Application_) SetWindowState(rhs int32)  {
	_ = this.PropPut(0x0000005b, []interface{}{rhs})
}

func (this *Application_) DisplayAutoCompleteTips() bool {
	retVal, _ := this.PropGet(0x0000005c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayAutoCompleteTips(rhs bool)  {
	_ = this.PropPut(0x0000005c, []interface{}{rhs})
}

func (this *Application_) Options() *Options {
	retVal, _ := this.PropGet(0x0000005d, nil)
	return NewOptions(retVal.IDispatch(), false, true)
}

func (this *Application_) DisplayAlerts() int32 {
	retVal, _ := this.PropGet(0x0000005e, nil)
	return retVal.LValVal()
}

func (this *Application_) SetDisplayAlerts(rhs int32)  {
	_ = this.PropPut(0x0000005e, []interface{}{rhs})
}

func (this *Application_) CustomDictionaries() *Dictionaries {
	retVal, _ := this.PropGet(0x0000005f, nil)
	return NewDictionaries(retVal.IDispatch(), false, true)
}

func (this *Application_) PathSeparator() string {
	retVal, _ := this.PropGet(0x00000060, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetStatusBar(rhs string)  {
	_ = this.PropPut(0x00000061, []interface{}{rhs})
}

func (this *Application_) MAPIAvailable() bool {
	retVal, _ := this.PropGet(0x00000062, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) DisplayScreenTips() bool {
	retVal, _ := this.PropGet(0x00000063, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayScreenTips(rhs bool)  {
	_ = this.PropPut(0x00000063, []interface{}{rhs})
}

func (this *Application_) EnableCancelKey() int32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *Application_) SetEnableCancelKey(rhs int32)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *Application_) UserControl() bool {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) FileSearch() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000067, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) MailSystem() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *Application_) DefaultTableSeparator() string {
	retVal, _ := this.PropGet(0x00000069, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetDefaultTableSeparator(rhs string)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *Application_) ShowVisualBasicEditor() bool {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowVisualBasicEditor(rhs bool)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *Application_) BrowseExtraFileTypes() string {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetBrowseExtraFileTypes(rhs string)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *Application_) IsObjectValid(object *win32.IUnknown) bool {
	retVal, _ := this.PropGet(0x0000006d, []interface{}{object})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) HangulHanjaDictionaries() *HangulHanjaConversionDictionaries {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return NewHangulHanjaConversionDictionaries(retVal.IDispatch(), false, true)
}

func (this *Application_) MailMessage() *MailMessage {
	retVal, _ := this.PropGet(0x0000015c, nil)
	return NewMailMessage(retVal.IDispatch(), false, true)
}

func (this *Application_) FocusInMailHeader() bool {
	retVal, _ := this.PropGet(0x00000182, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Application__Quit_OptArgs= []string{
	"SaveChanges", "OriginalFormat", "RouteDocument", 
}

func (this *Application_) Quit(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Application__Quit_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000451, nil, optArgs...)
	_= retVal
}

func (this *Application_) ScreenRefresh()  {
	retVal, _ := this.Call(0x0000012d, nil)
	_= retVal
}

var Application__PrintOutOld_OptArgs= []string{
	"Background", "Append", "Range", "OutputFileName", 
	"From", "To", "Item", "Copies", 
	"Pages", "PageType", "PrintToFile", "Collate", 
	"FileName", "ActivePrinterMacGX", "ManualDuplexPrint", 
}

func (this *Application_) PrintOutOld(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Application__PrintOutOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000012e, nil, optArgs...)
	_= retVal
}

func (this *Application_) LookupNameProperties(name string)  {
	retVal, _ := this.Call(0x0000012f, []interface{}{name})
	_= retVal
}

func (this *Application_) SubstituteFont(unavailableFont string, substituteFont string)  {
	retVal, _ := this.Call(0x00000130, []interface{}{unavailableFont, substituteFont})
	_= retVal
}

var Application__Repeat_OptArgs= []string{
	"Times", 
}

func (this *Application_) Repeat(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Application__Repeat_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000131, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) DDEExecute(channel int32, command string)  {
	retVal, _ := this.Call(0x00000136, []interface{}{channel, command})
	_= retVal
}

func (this *Application_) DDEInitiate(app string, topic string) int32 {
	retVal, _ := this.Call(0x00000137, []interface{}{app, topic})
	return retVal.LValVal()
}

func (this *Application_) DDEPoke(channel int32, item string, data string)  {
	retVal, _ := this.Call(0x00000138, []interface{}{channel, item, data})
	_= retVal
}

func (this *Application_) DDERequest(channel int32, item string) string {
	retVal, _ := this.Call(0x00000139, []interface{}{channel, item})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) DDETerminate(channel int32)  {
	retVal, _ := this.Call(0x0000013a, []interface{}{channel})
	_= retVal
}

func (this *Application_) DDETerminateAll()  {
	retVal, _ := this.Call(0x0000013b, nil)
	_= retVal
}

var Application__BuildKeyCode_OptArgs= []string{
	"Arg2", "Arg3", "Arg4", 
}

func (this *Application_) BuildKeyCode(arg1 int32, optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Application__BuildKeyCode_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000013c, []interface{}{arg1}, optArgs...)
	return retVal.LValVal()
}

var Application__KeyString_OptArgs= []string{
	"KeyCode2", 
}

func (this *Application_) KeyString(keyCode int32, optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Application__KeyString_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000013d, []interface{}{keyCode}, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) OrganizerCopy(source string, destination string, name string, object int32)  {
	retVal, _ := this.Call(0x0000013e, []interface{}{source, destination, name, object})
	_= retVal
}

func (this *Application_) OrganizerDelete(source string, name string, object int32)  {
	retVal, _ := this.Call(0x0000013f, []interface{}{source, name, object})
	_= retVal
}

func (this *Application_) OrganizerRename(source string, name string, newName string, object int32)  {
	retVal, _ := this.Call(0x00000140, []interface{}{source, name, newName, object})
	_= retVal
}

func (this *Application_) AddAddress(tagID **win32.SAFEARRAY, value **win32.SAFEARRAY)  {
	retVal, _ := this.Call(0x00000141, []interface{}{tagID, value})
	_= retVal
}

var Application__GetAddress_OptArgs= []string{
	"Name", "AddressProperties", "UseAutoText", "DisplaySelectDialog", 
	"SelectDialog", "CheckNamesDialog", "RecentAddressesChoice", "UpdateRecentAddresses", 
}

func (this *Application_) GetAddress(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(Application__GetAddress_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000142, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) CheckGrammar(string string) bool {
	retVal, _ := this.Call(0x00000143, []interface{}{string})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Application__CheckSpelling_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "MainDictionary", "CustomDictionary2", 
	"CustomDictionary3", "CustomDictionary4", "CustomDictionary5", "CustomDictionary6", 
	"CustomDictionary7", "CustomDictionary8", "CustomDictionary9", "CustomDictionary10", 
}

func (this *Application_) CheckSpelling(word string, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Application__CheckSpelling_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000144, []interface{}{word}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) ResetIgnoreAll()  {
	retVal, _ := this.Call(0x00000146, nil)
	_= retVal
}

var Application__GetSpellingSuggestions_OptArgs= []string{
	"CustomDictionary", "IgnoreUppercase", "MainDictionary", "SuggestionMode", 
	"CustomDictionary2", "CustomDictionary3", "CustomDictionary4", "CustomDictionary5", 
	"CustomDictionary6", "CustomDictionary7", "CustomDictionary8", "CustomDictionary9", "CustomDictionary10", 
}

func (this *Application_) GetSpellingSuggestions(word string, optArgs ...interface{}) *SpellingSuggestions {
	optArgs = ole.ProcessOptArgs(Application__GetSpellingSuggestions_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000147, []interface{}{word}, optArgs...)
	return NewSpellingSuggestions(retVal.IDispatch(), false, true)
}

func (this *Application_) GoBack()  {
	retVal, _ := this.Call(0x00000148, nil)
	_= retVal
}

func (this *Application_) Help(helpType *ole.Variant)  {
	retVal, _ := this.Call(0x00000149, []interface{}{helpType})
	_= retVal
}

func (this *Application_) AutomaticChange()  {
	retVal, _ := this.Call(0x0000014a, nil)
	_= retVal
}

func (this *Application_) ShowMe()  {
	retVal, _ := this.Call(0x0000014b, nil)
	_= retVal
}

func (this *Application_) HelpTool()  {
	retVal, _ := this.Call(0x0000014c, nil)
	_= retVal
}

func (this *Application_) NewWindow() *Window {
	retVal, _ := this.Call(0x00000159, nil)
	return NewWindow(retVal.IDispatch(), false, true)
}

func (this *Application_) ListCommands(listAllCommands bool)  {
	retVal, _ := this.Call(0x0000015a, []interface{}{listAllCommands})
	_= retVal
}

func (this *Application_) ShowClipboard()  {
	retVal, _ := this.Call(0x0000015d, nil)
	_= retVal
}

var Application__OnTime_OptArgs= []string{
	"Tolerance", 
}

func (this *Application_) OnTime(when *ole.Variant, name string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Application__OnTime_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000015e, []interface{}{when, name}, optArgs...)
	_= retVal
}

func (this *Application_) NextLetter()  {
	retVal, _ := this.Call(0x0000015f, nil)
	_= retVal
}

var Application__MountVolume_OptArgs= []string{
	"User", "UserPassword", "VolumePassword", 
}

func (this *Application_) MountVolume(zone string, server string, volume string, optArgs ...interface{}) int16 {
	optArgs = ole.ProcessOptArgs(Application__MountVolume_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000161, []interface{}{zone, server, volume}, optArgs...)
	return retVal.IValVal()
}

func (this *Application_) CleanString(string string) string {
	retVal, _ := this.Call(0x00000162, []interface{}{string})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SendFax()  {
	retVal, _ := this.Call(0x00000164, nil)
	_= retVal
}

func (this *Application_) ChangeFileOpenDirectory(path string)  {
	retVal, _ := this.Call(0x00000165, []interface{}{path})
	_= retVal
}

func (this *Application_) RunOld(macroName string)  {
	retVal, _ := this.Call(0x00000166, []interface{}{macroName})
	_= retVal
}

func (this *Application_) GoForward()  {
	retVal, _ := this.Call(0x00000167, nil)
	_= retVal
}

func (this *Application_) Move(left int32, top int32)  {
	retVal, _ := this.Call(0x00000168, []interface{}{left, top})
	_= retVal
}

func (this *Application_) Resize(width int32, height int32)  {
	retVal, _ := this.Call(0x00000169, []interface{}{width, height})
	_= retVal
}

func (this *Application_) InchesToPoints(inches float32) float32 {
	retVal, _ := this.Call(0x00000172, []interface{}{inches})
	return retVal.FltValVal()
}

func (this *Application_) CentimetersToPoints(centimeters float32) float32 {
	retVal, _ := this.Call(0x00000173, []interface{}{centimeters})
	return retVal.FltValVal()
}

func (this *Application_) MillimetersToPoints(millimeters float32) float32 {
	retVal, _ := this.Call(0x00000174, []interface{}{millimeters})
	return retVal.FltValVal()
}

func (this *Application_) PicasToPoints(picas float32) float32 {
	retVal, _ := this.Call(0x00000175, []interface{}{picas})
	return retVal.FltValVal()
}

func (this *Application_) LinesToPoints(lines float32) float32 {
	retVal, _ := this.Call(0x00000176, []interface{}{lines})
	return retVal.FltValVal()
}

func (this *Application_) PointsToInches(points float32) float32 {
	retVal, _ := this.Call(0x0000017c, []interface{}{points})
	return retVal.FltValVal()
}

func (this *Application_) PointsToCentimeters(points float32) float32 {
	retVal, _ := this.Call(0x0000017d, []interface{}{points})
	return retVal.FltValVal()
}

func (this *Application_) PointsToMillimeters(points float32) float32 {
	retVal, _ := this.Call(0x0000017e, []interface{}{points})
	return retVal.FltValVal()
}

func (this *Application_) PointsToPicas(points float32) float32 {
	retVal, _ := this.Call(0x0000017f, []interface{}{points})
	return retVal.FltValVal()
}

func (this *Application_) PointsToLines(points float32) float32 {
	retVal, _ := this.Call(0x00000180, []interface{}{points})
	return retVal.FltValVal()
}

func (this *Application_) Activate()  {
	retVal, _ := this.Call(0x00000181, nil)
	_= retVal
}

var Application__PointsToPixels_OptArgs= []string{
	"fVertical", 
}

func (this *Application_) PointsToPixels(points float32, optArgs ...interface{}) float32 {
	optArgs = ole.ProcessOptArgs(Application__PointsToPixels_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000183, []interface{}{points}, optArgs...)
	return retVal.FltValVal()
}

var Application__PixelsToPoints_OptArgs= []string{
	"fVertical", 
}

func (this *Application_) PixelsToPoints(pixels float32, optArgs ...interface{}) float32 {
	optArgs = ole.ProcessOptArgs(Application__PixelsToPoints_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000184, []interface{}{pixels}, optArgs...)
	return retVal.FltValVal()
}

func (this *Application_) KeyboardLatin()  {
	retVal, _ := this.Call(0x00000190, nil)
	_= retVal
}

func (this *Application_) KeyboardBidi()  {
	retVal, _ := this.Call(0x00000191, nil)
	_= retVal
}

func (this *Application_) ToggleKeyboard()  {
	retVal, _ := this.Call(0x00000192, nil)
	_= retVal
}

var Application__Keyboard_OptArgs= []string{
	"LangId", 
}

func (this *Application_) Keyboard(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(Application__Keyboard_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001be, nil, optArgs...)
	return retVal.LValVal()
}

func (this *Application_) ProductCode() string {
	retVal, _ := this.Call(0x00000194, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) DefaultWebOptions() *DefaultWebOptions {
	retVal, _ := this.Call(0x00000195, nil)
	return NewDefaultWebOptions(retVal.IDispatch(), false, true)
}

func (this *Application_) DiscussionSupport(range_ *ole.Variant, cid *ole.Variant, piCSE *ole.Variant)  {
	retVal, _ := this.Call(0x00000197, []interface{}{range_, cid, piCSE})
	_= retVal
}

func (this *Application_) SetDefaultTheme(name string, documentType int32)  {
	retVal, _ := this.Call(0x0000019e, []interface{}{name, documentType})
	_= retVal
}

func (this *Application_) GetDefaultTheme(documentType int32) string {
	retVal, _ := this.Call(0x000001a0, []interface{}{documentType})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) EmailOptions() *EmailOptions {
	retVal, _ := this.PropGet(0x00000185, nil)
	return NewEmailOptions(retVal.IDispatch(), false, true)
}

func (this *Application_) Language() int32 {
	retVal, _ := this.PropGet(0x00000187, nil)
	return retVal.LValVal()
}

func (this *Application_) COMAddIns() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) CheckLanguage() bool {
	retVal, _ := this.PropGet(0x00000070, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetCheckLanguage(rhs bool)  {
	_ = this.PropPut(0x00000070, []interface{}{rhs})
}

func (this *Application_) LanguageSettings() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000193, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) Dummy1() bool {
	retVal, _ := this.PropGet(0x00000196, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) AnswerWizard() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000199, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) FeatureInstall() int32 {
	retVal, _ := this.PropGet(0x000001bf, nil)
	return retVal.LValVal()
}

func (this *Application_) SetFeatureInstall(rhs int32)  {
	_ = this.PropPut(0x000001bf, []interface{}{rhs})
}

var Application__PrintOut2000_OptArgs= []string{
	"Background", "Append", "Range", "OutputFileName", 
	"From", "To", "Item", "Copies", 
	"Pages", "PageType", "PrintToFile", "Collate", 
	"FileName", "ActivePrinterMacGX", "ManualDuplexPrint", "PrintZoomColumn", 
	"PrintZoomRow", "PrintZoomPaperWidth", "PrintZoomPaperHeight", 
}

func (this *Application_) PrintOut2000(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Application__PrintOut2000_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bc, nil, optArgs...)
	_= retVal
}

var Application__Run_OptArgs= []string{
	"varg1", "varg2", "varg3", "varg4", 
	"varg5", "varg6", "varg7", "varg8", 
	"varg9", "varg10", "varg11", "varg12", 
	"varg13", "varg14", "varg15", "varg16", 
	"varg17", "varg18", "varg19", "varg20", 
	"varg21", "varg22", "varg23", "varg24", 
	"varg25", "varg26", "varg27", "varg28", 
	"varg29", "varg30", 
}

func (this *Application_) Run(macroName string, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Application__Run_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bd, []interface{}{macroName}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

var Application__PrintOut_OptArgs= []string{
	"Background", "Append", "Range", "OutputFileName", 
	"From", "To", "Item", "Copies", 
	"Pages", "PageType", "PrintToFile", "Collate", 
	"FileName", "ActivePrinterMacGX", "ManualDuplexPrint", "PrintZoomColumn", 
	"PrintZoomRow", "PrintZoomPaperWidth", "PrintZoomPaperHeight", 
}

func (this *Application_) PrintOut(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Application__PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001c0, nil, optArgs...)
	_= retVal
}

func (this *Application_) AutomationSecurity() int32 {
	retVal, _ := this.PropGet(0x000001c1, nil)
	return retVal.LValVal()
}

func (this *Application_) SetAutomationSecurity(rhs int32)  {
	_ = this.PropPut(0x000001c1, []interface{}{rhs})
}

func (this *Application_) FileDialog(fileDialogType int32) *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001c2, []interface{}{fileDialogType})
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) EmailTemplate() string {
	retVal, _ := this.PropGet(0x000001c3, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) SetEmailTemplate(rhs string)  {
	_ = this.PropPut(0x000001c3, []interface{}{rhs})
}

func (this *Application_) ShowWindowsInTaskbar() bool {
	retVal, _ := this.PropGet(0x000001c4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowWindowsInTaskbar(rhs bool)  {
	_ = this.PropPut(0x000001c4, []interface{}{rhs})
}

func (this *Application_) NewDocument() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001c6, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) ShowStartupDialog() bool {
	retVal, _ := this.PropGet(0x000001c7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowStartupDialog(rhs bool)  {
	_ = this.PropPut(0x000001c7, []interface{}{rhs})
}

func (this *Application_) AutoCorrectEmail() *AutoCorrect {
	retVal, _ := this.PropGet(0x000001c8, nil)
	return NewAutoCorrect(retVal.IDispatch(), false, true)
}

func (this *Application_) TaskPanes() *TaskPanes {
	retVal, _ := this.PropGet(0x000001c9, nil)
	return NewTaskPanes(retVal.IDispatch(), false, true)
}

func (this *Application_) DefaultLegalBlackline() bool {
	retVal, _ := this.PropGet(0x000001cb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDefaultLegalBlackline(rhs bool)  {
	_ = this.PropPut(0x000001cb, []interface{}{rhs})
}

func (this *Application_) Dummy2() bool {
	retVal, _ := this.Call(0x000001ca, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SmartTagRecognizers() *SmartTagRecognizers {
	retVal, _ := this.PropGet(0x000001cc, nil)
	return NewSmartTagRecognizers(retVal.IDispatch(), false, true)
}

func (this *Application_) SmartTagTypes() *SmartTagTypes {
	retVal, _ := this.PropGet(0x000001cd, nil)
	return NewSmartTagTypes(retVal.IDispatch(), false, true)
}

func (this *Application_) XMLNamespaces() *XMLNamespaces {
	retVal, _ := this.PropGet(0x000001cf, nil)
	return NewXMLNamespaces(retVal.IDispatch(), false, true)
}

func (this *Application_) PutFocusInMailHeader()  {
	retVal, _ := this.Call(0x000001d0, nil)
	_= retVal
}

func (this *Application_) ArbitraryXMLSupportAvailable() bool {
	retVal, _ := this.PropGet(0x000001d1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) BuildFull() string {
	retVal, _ := this.PropGet(0x000001d2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) BuildFeatureCrew() string {
	retVal, _ := this.PropGet(0x000001d3, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Application_) LoadMasterList(fileName string)  {
	retVal, _ := this.Call(0x000001d5, []interface{}{fileName})
	_= retVal
}

var Application__CompareDocuments_OptArgs= []string{
	"Destination", "Granularity", "CompareFormatting", "CompareCaseChanges", 
	"CompareWhitespace", "CompareTables", "CompareHeaders", "CompareFootnotes", 
	"CompareTextboxes", "CompareFields", "CompareComments", "CompareMoves", 
	"RevisedAuthor", "IgnoreAllComparisonWarnings", 
}

func (this *Application_) CompareDocuments(originalDocument *Document, revisedDocument *Document, optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(Application__CompareDocuments_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001d6, []interface{}{originalDocument, revisedDocument}, optArgs...)
	return NewDocument(retVal.IDispatch(), false, true)
}

var Application__MergeDocuments_OptArgs= []string{
	"Destination", "Granularity", "CompareFormatting", "CompareCaseChanges", 
	"CompareWhitespace", "CompareTables", "CompareHeaders", "CompareFootnotes", 
	"CompareTextboxes", "CompareFields", "CompareComments", "CompareMoves", 
	"OriginalAuthor", "RevisedAuthor", "FormatFrom", 
}

func (this *Application_) MergeDocuments(originalDocument *Document, revisedDocument *Document, optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(Application__MergeDocuments_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001d7, []interface{}{originalDocument, revisedDocument}, optArgs...)
	return NewDocument(retVal.IDispatch(), false, true)
}

func (this *Application_) Bibliography() *Bibliography {
	retVal, _ := this.PropGet(0x000001d8, nil)
	return NewBibliography(retVal.IDispatch(), false, true)
}

func (this *Application_) ShowStylePreviews() bool {
	retVal, _ := this.PropGet(0x000001d9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetShowStylePreviews(rhs bool)  {
	_ = this.PropPut(0x000001d9, []interface{}{rhs})
}

func (this *Application_) RestrictLinkedStyles() bool {
	retVal, _ := this.PropGet(0x000001da, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetRestrictLinkedStyles(rhs bool)  {
	_ = this.PropPut(0x000001da, []interface{}{rhs})
}

func (this *Application_) OMathAutoCorrect() *OMathAutoCorrect {
	retVal, _ := this.PropGet(0x000001db, nil)
	return NewOMathAutoCorrect(retVal.IDispatch(), false, true)
}

func (this *Application_) DisplayDocumentInformationPanel() bool {
	retVal, _ := this.PropGet(0x000001dc, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDisplayDocumentInformationPanel(rhs bool)  {
	_ = this.PropPut(0x000001dc, []interface{}{rhs})
}

func (this *Application_) Assistance() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001dd, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) OpenAttachmentsInFullScreen() bool {
	retVal, _ := this.PropGet(0x000001de, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetOpenAttachmentsInFullScreen(rhs bool)  {
	_ = this.PropPut(0x000001de, []interface{}{rhs})
}

func (this *Application_) ActiveEncryptionSession() int32 {
	retVal, _ := this.PropGet(0x000001df, nil)
	return retVal.LValVal()
}

func (this *Application_) DontResetInsertionPointProperties() bool {
	retVal, _ := this.PropGet(0x000001e0, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) SetDontResetInsertionPointProperties(rhs bool)  {
	_ = this.PropPut(0x000001e0, []interface{}{rhs})
}

func (this *Application_) SmartArtLayouts() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001e1, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) SmartArtQuickStyles() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001e2, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) SmartArtColors() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001e3, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) ThreeWayMerge(localDocument *Document, serverDocument *Document, baseDocument *Document, favorSource bool)  {
	retVal, _ := this.Call(0x000001e4, []interface{}{localDocument, serverDocument, baseDocument, favorSource})
	_= retVal
}

func (this *Application_) Dummy4()  {
	retVal, _ := this.Call(0x000001e5, nil)
	_= retVal
}

func (this *Application_) UndoRecord() *UndoRecord {
	retVal, _ := this.PropGet(0x000001e6, nil)
	return NewUndoRecord(retVal.IDispatch(), false, true)
}

func (this *Application_) PickerDialog() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000001e9, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Application_) ProtectedViewWindows() *ProtectedViewWindows {
	retVal, _ := this.PropGet(0x000001ea, nil)
	return NewProtectedViewWindows(retVal.IDispatch(), false, true)
}

func (this *Application_) ActiveProtectedViewWindow() *ProtectedViewWindow {
	retVal, _ := this.PropGet(0x000001eb, nil)
	return NewProtectedViewWindow(retVal.IDispatch(), false, true)
}

func (this *Application_) IsSandboxed() bool {
	retVal, _ := this.PropGet(0x000001ec, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Application_) FileValidation() int32 {
	retVal, _ := this.PropGet(0x000001ed, nil)
	return retVal.LValVal()
}

func (this *Application_) SetFileValidation(rhs int32)  {
	_ = this.PropPut(0x000001ed, []interface{}{rhs})
}

