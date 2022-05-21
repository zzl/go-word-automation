package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020911-0000-0000-C000-000000000046
var IID_TableOfAuthorities = syscall.GUID{0x00020911, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TableOfAuthorities struct {
	ole.OleClient
}

func NewTableOfAuthorities(pDisp *win32.IDispatch, addRef bool, scoped bool) *TableOfAuthorities {
	p := &TableOfAuthorities{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TableOfAuthoritiesFromVar(v ole.Variant) *TableOfAuthorities {
	return NewTableOfAuthorities(v.PdispValVal(), false, false)
}

func (this *TableOfAuthorities) IID() *syscall.GUID {
	return &IID_TableOfAuthorities
}

func (this *TableOfAuthorities) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TableOfAuthorities) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *TableOfAuthorities) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TableOfAuthorities) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *TableOfAuthorities) Passim() bool {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfAuthorities) SetPassim(rhs bool)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) KeepEntryFormatting() bool {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfAuthorities) SetKeepEntryFormatting(rhs bool)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) Category() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *TableOfAuthorities) SetCategory(rhs int32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) Bookmark() string {
	retVal := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfAuthorities) SetBookmark(rhs string)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) Separator() string {
	retVal := this.PropGet(0x00000005, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfAuthorities) SetSeparator(rhs string)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) IncludeSequenceName() string {
	retVal := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfAuthorities) SetIncludeSequenceName(rhs string)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) EntrySeparator() string {
	retVal := this.PropGet(0x00000007, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfAuthorities) SetEntrySeparator(rhs string)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) PageRangeSeparator() string {
	retVal := this.PropGet(0x00000008, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfAuthorities) SetPageRangeSeparator(rhs string)  {
	retVal := this.PropPut(0x00000008, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) IncludeCategoryHeader() bool {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TableOfAuthorities) SetIncludeCategoryHeader(rhs bool)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) PageNumberSeparator() string {
	retVal := this.PropGet(0x0000000a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfAuthorities) SetPageNumberSeparator(rhs string)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) Range() *Range {
	retVal := this.PropGet(0x0000000b, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *TableOfAuthorities) TabLeader() int32 {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *TableOfAuthorities) SetTabLeader(rhs int32)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *TableOfAuthorities) Delete()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

func (this *TableOfAuthorities) Update()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

