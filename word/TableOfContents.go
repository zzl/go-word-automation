package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020913-0000-0000-C000-000000000046
var IID_TableOfContents = syscall.GUID{0x00020913, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TableOfContents struct {
	ole.OleClient
}

func NewTableOfContents(pDisp *win32.IDispatch, addRef bool, scoped bool) *TableOfContents {
	 if pDisp == nil {
		return nil;
	}
	p := &TableOfContents{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TableOfContentsFromVar(v ole.Variant) *TableOfContents {
	return NewTableOfContents(v.IDispatch(), false, false)
}

func (this *TableOfContents) IID() *syscall.GUID {
	return &IID_TableOfContents
}

func (this *TableOfContents) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TableOfContents) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TableOfContents) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TableOfContents) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TableOfContents) UseHeadingStyles() bool {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfContents) SetUseHeadingStyles(rhs bool)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *TableOfContents) UseFields() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfContents) SetUseFields(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *TableOfContents) UpperHeadingLevel() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *TableOfContents) SetUpperHeadingLevel(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *TableOfContents) LowerHeadingLevel() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *TableOfContents) SetLowerHeadingLevel(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *TableOfContents) TableID() string {
	retVal, _ := this.PropGet(0x00000005, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfContents) SetTableID(rhs string)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *TableOfContents) HeadingStyles() *HeadingStyles {
	retVal, _ := this.PropGet(0x00000006, nil)
	return NewHeadingStyles(retVal.IDispatch(), false, true)
}

func (this *TableOfContents) RightAlignPageNumbers() bool {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfContents) SetRightAlignPageNumbers(rhs bool)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *TableOfContents) IncludePageNumbers() bool {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfContents) SetIncludePageNumbers(rhs bool)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *TableOfContents) Range() *Range {
	retVal, _ := this.PropGet(0x00000009, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *TableOfContents) TabLeader() int32 {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *TableOfContents) SetTabLeader(rhs int32)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *TableOfContents) Delete()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

func (this *TableOfContents) UpdatePageNumbers()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *TableOfContents) Update()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *TableOfContents) UseHyperlinks() bool {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfContents) SetUseHyperlinks(rhs bool)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *TableOfContents) HidePageNumbersInWeb() bool {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfContents) SetHidePageNumbersInWeb(rhs bool)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

