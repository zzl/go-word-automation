package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020986-0000-0000-C000-000000000046
var IID_PageNumbers = syscall.GUID{0x00020986, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PageNumbers struct {
	ole.OleClient
}

func NewPageNumbers(pDisp *win32.IDispatch, addRef bool, scoped bool) *PageNumbers {
	 if pDisp == nil {
		return nil;
	}
	p := &PageNumbers{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PageNumbersFromVar(v ole.Variant) *PageNumbers {
	return NewPageNumbers(v.IDispatch(), false, false)
}

func (this *PageNumbers) IID() *syscall.GUID {
	return &IID_PageNumbers
}

func (this *PageNumbers) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PageNumbers) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PageNumbers) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *PageNumbers) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PageNumbers) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *PageNumbers) ForEach(action func(item *PageNumber) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*PageNumber)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *PageNumbers) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *PageNumbers) NumberStyle() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *PageNumbers) SetNumberStyle(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *PageNumbers) IncludeChapterNumber() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageNumbers) SetIncludeChapterNumber(rhs bool)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *PageNumbers) HeadingLevelForChapter() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *PageNumbers) SetHeadingLevelForChapter(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *PageNumbers) ChapterPageSeparator() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *PageNumbers) SetChapterPageSeparator(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *PageNumbers) RestartNumberingAtSection() bool {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageNumbers) SetRestartNumberingAtSection(rhs bool)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *PageNumbers) StartingNumber() int32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *PageNumbers) SetStartingNumber(rhs int32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *PageNumbers) ShowFirstPageNumber() bool {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageNumbers) SetShowFirstPageNumber(rhs bool)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *PageNumbers) Item(index int32) *PageNumber {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewPageNumber(retVal.IDispatch(), false, true)
}

var PageNumbers_Add_OptArgs= []string{
	"PageNumberAlignment", "FirstPage", 
}

func (this *PageNumbers) Add(optArgs ...interface{}) *PageNumber {
	optArgs = ole.ProcessOptArgs(PageNumbers_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, nil, optArgs...)
	return NewPageNumber(retVal.IDispatch(), false, true)
}

func (this *PageNumbers) DoubleQuote() bool {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *PageNumbers) SetDoubleQuote(rhs bool)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

