package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// EE95AFE3-3026-4172-B078-0E79DAB5CC3D
var IID_ContentControl = syscall.GUID{0xEE95AFE3, 0x3026, 0x4172, 
	[8]byte{0xB0, 0x78, 0x0E, 0x79, 0xDA, 0xB5, 0xCC, 0x3D}}

type ContentControl struct {
	ole.OleClient
}

func NewContentControl(pDisp *win32.IDispatch, addRef bool, scoped bool) *ContentControl {
	 if pDisp == nil {
		return nil;
	}
	p := &ContentControl{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ContentControlFromVar(v ole.Variant) *ContentControl {
	return NewContentControl(v.IDispatch(), false, false)
}

func (this *ContentControl) IID() *syscall.GUID {
	return &IID_ContentControl
}

func (this *ContentControl) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ContentControl) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ContentControl) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ContentControl) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ContentControl) Range() *Range {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *ContentControl) LockContentControl() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetLockContentControl(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *ContentControl) LockContents() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetLockContents(rhs bool)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *ContentControl) XMLMapping() *XMLMapping {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewXMLMapping(retVal.IDispatch(), false, true)
}

func (this *ContentControl) Type() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetType(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *ContentControl) Copy()  {
	retVal, _ := this.Call(0x00000006, nil)
	_= retVal
}

func (this *ContentControl) Cut()  {
	retVal, _ := this.Call(0x00000007, nil)
	_= retVal
}

var ContentControl_Delete_OptArgs= []string{
	"DeleteContents", 
}

func (this *ContentControl) Delete(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ContentControl_Delete_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000008, nil, optArgs...)
	_= retVal
}

func (this *ContentControl) DropdownListEntries() *ContentControlListEntries {
	retVal, _ := this.PropGet(0x00000009, nil)
	return NewContentControlListEntries(retVal.IDispatch(), false, true)
}

func (this *ContentControl) PlaceholderText() *BuildingBlock {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewBuildingBlock(retVal.IDispatch(), false, true)
}

var ContentControl_SetPlaceholderText_OptArgs= []string{
	"BuildingBlock", "Range", "Text", 
}

func (this *ContentControl) SetPlaceholderText(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ContentControl_SetPlaceholderText_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000b, nil, optArgs...)
	_= retVal
}

func (this *ContentControl) Title() string {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) SetTitle(rhs string)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *ContentControl) DateDisplayFormat() string {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) SetDateDisplayFormat(rhs string)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *ContentControl) MultiLine() bool {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetMultiLine(rhs bool)  {
	_ = this.PropPut(0x0000000f, []interface{}{rhs})
}

func (this *ContentControl) ParentContentControl() *ContentControl {
	retVal, _ := this.PropGet(0x00000010, nil)
	return NewContentControl(retVal.IDispatch(), false, true)
}

func (this *ContentControl) Temporary() bool {
	retVal, _ := this.PropGet(0x00000011, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetTemporary(rhs bool)  {
	_ = this.PropPut(0x00000011, []interface{}{rhs})
}

func (this *ContentControl) ID() string {
	retVal, _ := this.PropGet(0x00000012, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) ShowingPlaceholderText() bool {
	retVal, _ := this.PropGet(0x00000013, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) DateStorageFormat() int32 {
	retVal, _ := this.PropGet(0x00000014, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetDateStorageFormat(rhs int32)  {
	_ = this.PropPut(0x00000014, []interface{}{rhs})
}

func (this *ContentControl) BuildingBlockType() int32 {
	retVal, _ := this.PropGet(0x00000015, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetBuildingBlockType(rhs int32)  {
	_ = this.PropPut(0x00000015, []interface{}{rhs})
}

func (this *ContentControl) BuildingBlockCategory() string {
	retVal, _ := this.PropGet(0x00000016, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) SetBuildingBlockCategory(rhs string)  {
	_ = this.PropPut(0x00000016, []interface{}{rhs})
}

func (this *ContentControl) DateDisplayLocale() int32 {
	retVal, _ := this.PropGet(0x00000017, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetDateDisplayLocale(rhs int32)  {
	_ = this.PropPut(0x00000017, []interface{}{rhs})
}

func (this *ContentControl) Ungroup()  {
	retVal, _ := this.Call(0x00000018, nil)
	_= retVal
}

func (this *ContentControl) DefaultTextStyle() ole.Variant {
	retVal, _ := this.PropGet(0x00000019, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ContentControl) SetDefaultTextStyle(rhs *ole.Variant)  {
	_ = this.PropPut(0x00000019, []interface{}{rhs})
}

func (this *ContentControl) DateCalendarType() int32 {
	retVal, _ := this.PropGet(0x0000001a, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetDateCalendarType(rhs int32)  {
	_ = this.PropPut(0x0000001a, []interface{}{rhs})
}

func (this *ContentControl) Tag() string {
	retVal, _ := this.PropGet(0x0000001b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) SetTag(rhs string)  {
	_ = this.PropPut(0x0000001b, []interface{}{rhs})
}

func (this *ContentControl) Checked() bool {
	retVal, _ := this.PropGet(0x0000001c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetChecked(rhs bool)  {
	_ = this.PropPut(0x0000001c, []interface{}{rhs})
}

var ContentControl_SetCheckedSymbol_OptArgs= []string{
	"Font", 
}

func (this *ContentControl) SetCheckedSymbol(characterNumber int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ContentControl_SetCheckedSymbol_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000001d, []interface{}{characterNumber}, optArgs...)
	_= retVal
}

var ContentControl_SetUncheckedSymbol_OptArgs= []string{
	"Font", 
}

func (this *ContentControl) SetUncheckedSymbol(characterNumber int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ContentControl_SetUncheckedSymbol_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000001e, []interface{}{characterNumber}, optArgs...)
	_= retVal
}

