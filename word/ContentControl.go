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
	return NewContentControl(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ContentControl) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ContentControl) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ContentControl) Range() *Range {
	retVal := this.PropGet(0x00000001, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *ContentControl) LockContentControl() bool {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetLockContentControl(rhs bool)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) LockContents() bool {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetLockContents(rhs bool)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) XMLMapping() *XMLMapping {
	retVal := this.PropGet(0x00000004, nil)
	return NewXMLMapping(retVal.PdispValVal(), false, true)
}

func (this *ContentControl) Type() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetType(rhs int32)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) Copy()  {
	retVal := this.Call(0x00000006, nil)
	_= retVal
}

func (this *ContentControl) Cut()  {
	retVal := this.Call(0x00000007, nil)
	_= retVal
}

func (this *ContentControl) Delete(deleteContents bool)  {
	retVal := this.Call(0x00000008, []interface{}{deleteContents})
	_= retVal
}

func (this *ContentControl) DropdownListEntries() *ContentControlListEntries {
	retVal := this.PropGet(0x00000009, nil)
	return NewContentControlListEntries(retVal.PdispValVal(), false, true)
}

func (this *ContentControl) PlaceholderText() *BuildingBlock {
	retVal := this.PropGet(0x0000000a, nil)
	return NewBuildingBlock(retVal.PdispValVal(), false, true)
}

func (this *ContentControl) SetPlaceholderText(buildingBlock *BuildingBlock, range_ *Range, text string)  {
	retVal := this.Call(0x0000000b, []interface{}{buildingBlock, range_, text})
	_= retVal
}

func (this *ContentControl) Title() string {
	retVal := this.PropGet(0x0000000c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) SetTitle(rhs string)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) DateDisplayFormat() string {
	retVal := this.PropGet(0x0000000d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) SetDateDisplayFormat(rhs string)  {
	retVal := this.PropPut(0x0000000d, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) MultiLine() bool {
	retVal := this.PropGet(0x0000000f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetMultiLine(rhs bool)  {
	retVal := this.PropPut(0x0000000f, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) ParentContentControl() *ContentControl {
	retVal := this.PropGet(0x00000010, nil)
	return NewContentControl(retVal.PdispValVal(), false, true)
}

func (this *ContentControl) Temporary() bool {
	retVal := this.PropGet(0x00000011, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetTemporary(rhs bool)  {
	retVal := this.PropPut(0x00000011, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) ID() string {
	retVal := this.PropGet(0x00000012, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) ShowingPlaceholderText() bool {
	retVal := this.PropGet(0x00000013, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) DateStorageFormat() int32 {
	retVal := this.PropGet(0x00000014, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetDateStorageFormat(rhs int32)  {
	retVal := this.PropPut(0x00000014, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) BuildingBlockType() int32 {
	retVal := this.PropGet(0x00000015, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetBuildingBlockType(rhs int32)  {
	retVal := this.PropPut(0x00000015, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) BuildingBlockCategory() string {
	retVal := this.PropGet(0x00000016, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) SetBuildingBlockCategory(rhs string)  {
	retVal := this.PropPut(0x00000016, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) DateDisplayLocale() int32 {
	retVal := this.PropGet(0x00000017, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetDateDisplayLocale(rhs int32)  {
	retVal := this.PropPut(0x00000017, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) Ungroup()  {
	retVal := this.Call(0x00000018, nil)
	_= retVal
}

func (this *ContentControl) DefaultTextStyle() ole.Variant {
	retVal := this.PropGet(0x00000019, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ContentControl) SetDefaultTextStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000019, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) DateCalendarType() int32 {
	retVal := this.PropGet(0x0000001a, nil)
	return retVal.LValVal()
}

func (this *ContentControl) SetDateCalendarType(rhs int32)  {
	retVal := this.PropPut(0x0000001a, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) Tag() string {
	retVal := this.PropGet(0x0000001b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ContentControl) SetTag(rhs string)  {
	retVal := this.PropPut(0x0000001b, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) Checked() bool {
	retVal := this.PropGet(0x0000001c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ContentControl) SetChecked(rhs bool)  {
	retVal := this.PropPut(0x0000001c, []interface{}{rhs})
	_= retVal
}

func (this *ContentControl) SetCheckedSymbol(characterNumber int32, font string)  {
	retVal := this.Call(0x0000001d, []interface{}{characterNumber, font})
	_= retVal
}

func (this *ContentControl) SetUncheckedSymbol(characterNumber int32, font string)  {
	retVal := this.Call(0x0000001e, []interface{}{characterNumber, font})
	_= retVal
}

