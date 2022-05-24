package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002091D-0000-0000-C000-000000000046
var IID_MailMergeDataSource = syscall.GUID{0x0002091D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailMergeDataSource struct {
	ole.OleClient
}

func NewMailMergeDataSource(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailMergeDataSource {
	 if pDisp == nil {
		return nil;
	}
	p := &MailMergeDataSource{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailMergeDataSourceFromVar(v ole.Variant) *MailMergeDataSource {
	return NewMailMergeDataSource(v.IDispatch(), false, false)
}

func (this *MailMergeDataSource) IID() *syscall.GUID {
	return &IID_MailMergeDataSource
}

func (this *MailMergeDataSource) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailMergeDataSource) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *MailMergeDataSource) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataSource) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *MailMergeDataSource) Name() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMergeDataSource) HeaderSourceName() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMergeDataSource) Type() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataSource) HeaderSourceType() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataSource) ConnectString() string {
	retVal, _ := this.PropGet(0x00000005, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMergeDataSource) QueryString() string {
	retVal, _ := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMergeDataSource) SetQueryString(rhs string)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *MailMergeDataSource) ActiveRecord() int32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataSource) SetActiveRecord(rhs int32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *MailMergeDataSource) FirstRecord() int32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataSource) SetFirstRecord(rhs int32)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *MailMergeDataSource) LastRecord() int32 {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataSource) SetLastRecord(rhs int32)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *MailMergeDataSource) FieldNames() *MailMergeFieldNames {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewMailMergeFieldNames(retVal.IDispatch(), false, true)
}

func (this *MailMergeDataSource) DataFields() *MailMergeDataFields {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return NewMailMergeDataFields(retVal.IDispatch(), false, true)
}

func (this *MailMergeDataSource) FindRecord2000(findText string, field string) bool {
	retVal, _ := this.Call(0x00000065, []interface{}{findText, field})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailMergeDataSource) RecordCount() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataSource) Included() bool {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailMergeDataSource) SetIncluded(rhs bool)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *MailMergeDataSource) InvalidAddress() bool {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailMergeDataSource) SetInvalidAddress(rhs bool)  {
	_ = this.PropPut(0x0000000e, []interface{}{rhs})
}

func (this *MailMergeDataSource) InvalidComments() string {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MailMergeDataSource) SetInvalidComments(rhs string)  {
	_ = this.PropPut(0x0000000f, []interface{}{rhs})
}

func (this *MailMergeDataSource) MappedDataFields() *MappedDataFields {
	retVal, _ := this.PropGet(0x00000010, nil)
	return NewMappedDataFields(retVal.IDispatch(), false, true)
}

func (this *MailMergeDataSource) TableName() string {
	retVal, _ := this.PropGet(0x00000011, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var MailMergeDataSource_FindRecord_OptArgs= []string{
	"Field", 
}

func (this *MailMergeDataSource) FindRecord(findText string, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(MailMergeDataSource_FindRecord_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, []interface{}{findText}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *MailMergeDataSource) SetAllIncludedFlags(included bool)  {
	retVal, _ := this.Call(0x00000067, []interface{}{included})
	_= retVal
}

func (this *MailMergeDataSource) SetAllErrorFlags(invalid bool, invalidComment string)  {
	retVal, _ := this.Call(0x00000068, []interface{}{invalid, invalidComment})
	_= retVal
}

func (this *MailMergeDataSource) Close()  {
	retVal, _ := this.Call(0x00000069, nil)
	_= retVal
}

