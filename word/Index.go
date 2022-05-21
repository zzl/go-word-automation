package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002097D-0000-0000-C000-000000000046
var IID_Index = syscall.GUID{0x0002097D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Index struct {
	ole.OleClient
}

func NewIndex(pDisp *win32.IDispatch, addRef bool, scoped bool) *Index {
	p := &Index{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func IndexFromVar(v ole.Variant) *Index {
	return NewIndex(v.PdispValVal(), false, false)
}

func (this *Index) IID() *syscall.GUID {
	return &IID_Index
}

func (this *Index) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Index) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Index) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Index) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Index) HeadingSeparator() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Index) SetHeadingSeparator(rhs int32)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *Index) RightAlignPageNumbers() bool {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Index) SetRightAlignPageNumbers(rhs bool)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *Index) Type() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Index) SetType(rhs int32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *Index) NumberOfColumns() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Index) SetNumberOfColumns(rhs int32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *Index) Range() *Range {
	retVal := this.PropGet(0x00000005, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Index) TabLeader() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Index) SetTabLeader(rhs int32)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *Index) AccentedLetters() bool {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Index) SetAccentedLetters(rhs bool)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *Index) SortBy() int32 {
	retVal := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Index) SetSortBy(rhs int32)  {
	retVal := this.PropPut(0x00000008, []interface{}{rhs})
	_= retVal
}

func (this *Index) Filter() int32 {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.LValVal()
}

func (this *Index) SetFilter(rhs int32)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *Index) Delete()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

func (this *Index) Update()  {
	retVal := this.Call(0x00000066, nil)
	_= retVal
}

func (this *Index) IndexLanguage() int32 {
	retVal := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *Index) SetIndexLanguage(rhs int32)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

