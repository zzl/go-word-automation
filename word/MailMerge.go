package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020920-0000-0000-C000-000000000046
var IID_MailMerge = syscall.GUID{0x00020920, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailMerge struct {
	ole.OleClient
}

func NewMailMerge(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailMerge {
	p := &MailMerge{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailMergeFromVar(v ole.Variant) *MailMerge {
	return NewMailMerge(v.PdispValVal(), false, false)
}

func (this *MailMerge) IID() *syscall.GUID {
	return &IID_MailMerge
}

func (this *MailMerge) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailMerge) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MailMerge) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailMerge) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MailMerge) MainDocumentType() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *MailMerge) SetMainDocumentType(rhs int32)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *MailMerge) State() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *MailMerge) Destination() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *MailMerge) SetDestination(rhs int32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *MailMerge) DataSource() *MailMergeDataSource {
	retVal := this.PropGet(0x00000004, nil)
	return NewMailMergeDataSource(retVal.PdispValVal(), false, true)
}

func (this *MailMerge) Fields() *MailMergeFields {
	retVal := this.PropGet(0x00000005, nil)
	return NewMailMergeFields(retVal.PdispValVal(), false, true)
}

func (this *MailMerge) ViewMailMergeFieldCodes() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *MailMerge) SetViewMailMergeFieldCodes(rhs int32)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *MailMerge) SuppressBlankLines() bool {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailMerge) SetSuppressBlankLines(rhs bool)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *MailMerge) MailAsAttachment() bool {
	retVal := this.PropGet(0x00000008, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailMerge) SetMailAsAttachment(rhs bool)  {
	retVal := this.PropPut(0x00000008, []interface{}{rhs})
	_= retVal
}

func (this *MailMerge) MailAddressFieldName() string {
	retVal := this.PropGet(0x00000009, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMerge) SetMailAddressFieldName(rhs string)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *MailMerge) MailSubject() string {
	retVal := this.PropGet(0x0000000a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMerge) SetMailSubject(rhs string)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

var MailMerge_CreateDataSource_OptArgs= []string{
	"Name", "PasswordDocument", "WritePasswordDocument", "HeaderRecord", 
	"MSQuery", "SQLStatement", "SQLStatement1", "Connection", "LinkToSource", 
}

func (this *MailMerge) CreateDataSource(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailMerge_CreateDataSource_OptArgs, optArgs)
	retVal := this.Call(0x00000065, nil, optArgs...)
	_= retVal
}

var MailMerge_CreateHeaderSource_OptArgs= []string{
	"PasswordDocument", "WritePasswordDocument", "HeaderRecord", 
}

func (this *MailMerge) CreateHeaderSource(name string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailMerge_CreateHeaderSource_OptArgs, optArgs)
	retVal := this.Call(0x00000066, []interface{}{name}, optArgs...)
	_= retVal
}

var MailMerge_OpenDataSource2000_OptArgs= []string{
	"Format", "ConfirmConversions", "ReadOnly", "LinkToSource", 
	"AddToRecentFiles", "PasswordDocument", "PasswordTemplate", "Revert", 
	"WritePasswordDocument", "WritePasswordTemplate", "Connection", "SQLStatement", "SQLStatement1", 
}

func (this *MailMerge) OpenDataSource2000(name string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailMerge_OpenDataSource2000_OptArgs, optArgs)
	retVal := this.Call(0x00000067, []interface{}{name}, optArgs...)
	_= retVal
}

var MailMerge_OpenHeaderSource2000_OptArgs= []string{
	"Format", "ConfirmConversions", "ReadOnly", "AddToRecentFiles", 
	"PasswordDocument", "PasswordTemplate", "Revert", "WritePasswordDocument", "WritePasswordTemplate", 
}

func (this *MailMerge) OpenHeaderSource2000(name string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailMerge_OpenHeaderSource2000_OptArgs, optArgs)
	retVal := this.Call(0x00000068, []interface{}{name}, optArgs...)
	_= retVal
}

var MailMerge_Execute_OptArgs= []string{
	"Pause", 
}

func (this *MailMerge) Execute(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailMerge_Execute_OptArgs, optArgs)
	retVal := this.Call(0x00000069, nil, optArgs...)
	_= retVal
}

func (this *MailMerge) Check()  {
	retVal := this.Call(0x0000006a, nil)
	_= retVal
}

func (this *MailMerge) EditDataSource()  {
	retVal := this.Call(0x0000006b, nil)
	_= retVal
}

func (this *MailMerge) EditHeaderSource()  {
	retVal := this.Call(0x0000006c, nil)
	_= retVal
}

func (this *MailMerge) EditMainDocument()  {
	retVal := this.Call(0x0000006d, nil)
	_= retVal
}

func (this *MailMerge) UseAddressBook(type_ string)  {
	retVal := this.Call(0x0000006f, []interface{}{type_})
	_= retVal
}

func (this *MailMerge) HighlightMergeFields() bool {
	retVal := this.PropGet(0x0000000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailMerge) SetHighlightMergeFields(rhs bool)  {
	retVal := this.PropPut(0x0000000b, []interface{}{rhs})
	_= retVal
}

func (this *MailMerge) MailFormat() int32 {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *MailMerge) SetMailFormat(rhs int32)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *MailMerge) ShowSendToCustom() string {
	retVal := this.PropGet(0x0000000d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMerge) SetShowSendToCustom(rhs string)  {
	retVal := this.PropPut(0x0000000d, []interface{}{rhs})
	_= retVal
}

func (this *MailMerge) WizardState() int32 {
	retVal := this.PropGet(0x0000000e, nil)
	return retVal.LValVal()
}

func (this *MailMerge) SetWizardState(rhs int32)  {
	retVal := this.PropPut(0x0000000e, []interface{}{rhs})
	_= retVal
}

var MailMerge_OpenDataSource_OptArgs= []string{
	"Format", "ConfirmConversions", "ReadOnly", "LinkToSource", 
	"AddToRecentFiles", "PasswordDocument", "PasswordTemplate", "Revert", 
	"WritePasswordDocument", "WritePasswordTemplate", "Connection", "SQLStatement", 
	"SQLStatement1", "OpenExclusive", "SubType", 
}

func (this *MailMerge) OpenDataSource(name string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailMerge_OpenDataSource_OptArgs, optArgs)
	retVal := this.Call(0x00000070, []interface{}{name}, optArgs...)
	_= retVal
}

var MailMerge_OpenHeaderSource_OptArgs= []string{
	"Format", "ConfirmConversions", "ReadOnly", "AddToRecentFiles", 
	"PasswordDocument", "PasswordTemplate", "Revert", "WritePasswordDocument", 
	"WritePasswordTemplate", "OpenExclusive", 
}

func (this *MailMerge) OpenHeaderSource(name string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailMerge_OpenHeaderSource_OptArgs, optArgs)
	retVal := this.Call(0x00000071, []interface{}{name}, optArgs...)
	_= retVal
}

var MailMerge_ShowWizard_OptArgs= []string{
	"ShowDocumentStep", "ShowTemplateStep", "ShowDataStep", "ShowWriteStep", 
	"ShowPreviewStep", "ShowMergeStep", 
}

func (this *MailMerge) ShowWizard(initialState *ole.Variant, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailMerge_ShowWizard_OptArgs, optArgs)
	retVal := this.Call(0x00000072, []interface{}{initialState}, optArgs...)
	_= retVal
}

