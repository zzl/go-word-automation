package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020921-0000-0000-C000-000000000046
var IID_TableOfFigures = syscall.GUID{0x00020921, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TableOfFigures struct {
	ole.OleClient
}

func NewTableOfFigures(pDisp *win32.IDispatch, addRef bool, scoped bool) *TableOfFigures {
	 if pDisp == nil {
		return nil;
	}
	p := &TableOfFigures{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TableOfFiguresFromVar(v ole.Variant) *TableOfFigures {
	return NewTableOfFigures(v.IDispatch(), false, false)
}

func (this *TableOfFigures) IID() *syscall.GUID {
	return &IID_TableOfFigures
}

func (this *TableOfFigures) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TableOfFigures) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TableOfFigures) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TableOfFigures) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TableOfFigures) Caption() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfFigures) SetCaption(rhs string)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *TableOfFigures) IncludeLabel() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfFigures) SetIncludeLabel(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *TableOfFigures) RightAlignPageNumbers() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfFigures) SetRightAlignPageNumbers(rhs bool)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *TableOfFigures) UseHeadingStyles() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfFigures) SetUseHeadingStyles(rhs bool)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *TableOfFigures) LowerHeadingLevel() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *TableOfFigures) SetLowerHeadingLevel(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *TableOfFigures) UpperHeadingLevel() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *TableOfFigures) SetUpperHeadingLevel(rhs int32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *TableOfFigures) IncludePageNumbers() bool {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfFigures) SetIncludePageNumbers(rhs bool)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *TableOfFigures) Range() *Range {
	retVal, _ := this.PropGet(0x00000008, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *TableOfFigures) UseFields() bool {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfFigures) SetUseFields(rhs bool)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *TableOfFigures) TableID() string {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfFigures) SetTableID(rhs string)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *TableOfFigures) HeadingStyles() *HeadingStyles {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return NewHeadingStyles(retVal.IDispatch(), false, true)
}

func (this *TableOfFigures) TabLeader() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *TableOfFigures) SetTabLeader(rhs int32)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *TableOfFigures) Delete()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

func (this *TableOfFigures) UpdatePageNumbers()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *TableOfFigures) Update()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *TableOfFigures) UseHyperlinks() bool {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfFigures) SetUseHyperlinks(rhs bool)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *TableOfFigures) HidePageNumbersInWeb() bool {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfFigures) SetHidePageNumbersInWeb(rhs bool)  {
	_ = this.PropPut(0x0000000e, []interface{}{rhs})
}

