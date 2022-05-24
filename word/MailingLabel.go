package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020917-0000-0000-C000-000000000046
var IID_MailingLabel = syscall.GUID{0x00020917, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailingLabel struct {
	ole.OleClient
}

func NewMailingLabel(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailingLabel {
	 if pDisp == nil {
		return nil;
	}
	p := &MailingLabel{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailingLabelFromVar(v ole.Variant) *MailingLabel {
	return NewMailingLabel(v.IDispatch(), false, false)
}

func (this *MailingLabel) IID() *syscall.GUID {
	return &IID_MailingLabel
}

func (this *MailingLabel) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailingLabel) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *MailingLabel) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailingLabel) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *MailingLabel) DefaultPrintBarCode() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailingLabel) SetDefaultPrintBarCode(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *MailingLabel) DefaultLaserTray() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *MailingLabel) SetDefaultLaserTray(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *MailingLabel) CustomLabels() *CustomLabels {
	retVal, _ := this.PropGet(0x00000008, nil)
	return NewCustomLabels(retVal.IDispatch(), false, true)
}

func (this *MailingLabel) DefaultLabelName() string {
	retVal, _ := this.PropGet(0x00000009, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailingLabel) SetDefaultLabelName(rhs string)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

var MailingLabel_CreateNewDocument2000_OptArgs= []string{
	"Name", "Address", "AutoText", "ExtractAddress", "LaserTray", 
}

func (this *MailingLabel) CreateNewDocument2000(optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(MailingLabel_CreateNewDocument2000_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, nil, optArgs...)
	return NewDocument(retVal.IDispatch(), false, true)
}

var MailingLabel_PrintOut2000_OptArgs= []string{
	"Name", "Address", "ExtractAddress", "LaserTray", 
	"SingleLabel", "Row", "Column", 
}

func (this *MailingLabel) PrintOut2000(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailingLabel_PrintOut2000_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, nil, optArgs...)
	_= retVal
}

func (this *MailingLabel) LabelOptions()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

var MailingLabel_CreateNewDocument_OptArgs= []string{
	"Name", "Address", "AutoText", "ExtractAddress", 
	"LaserTray", "PrintEPostageLabel", "Vertical", 
}

func (this *MailingLabel) CreateNewDocument(optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(MailingLabel_CreateNewDocument_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000068, nil, optArgs...)
	return NewDocument(retVal.IDispatch(), false, true)
}

var MailingLabel_PrintOut_OptArgs= []string{
	"Name", "Address", "ExtractAddress", "LaserTray", 
	"SingleLabel", "Row", "Column", "PrintEPostageLabel", "Vertical", 
}

func (this *MailingLabel) PrintOut(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailingLabel_PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000069, nil, optArgs...)
	_= retVal
}

func (this *MailingLabel) Vertical() bool {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailingLabel) SetVertical(rhs bool)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

var MailingLabel_CreateNewDocumentByID_OptArgs= []string{
	"LabelID", "Address", "AutoText", "ExtractAddress", 
	"LaserTray", "PrintEPostageLabel", "Vertical", 
}

func (this *MailingLabel) CreateNewDocumentByID(optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(MailingLabel_CreateNewDocumentByID_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006a, nil, optArgs...)
	return NewDocument(retVal.IDispatch(), false, true)
}

var MailingLabel_PrintOutByID_OptArgs= []string{
	"LabelID", "Address", "ExtractAddress", "LaserTray", 
	"SingleLabel", "Row", "Column", "PrintEPostageLabel", "Vertical", 
}

func (this *MailingLabel) PrintOutByID(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(MailingLabel_PrintOutByID_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006b, nil, optArgs...)
	_= retVal
}

