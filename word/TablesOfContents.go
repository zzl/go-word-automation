package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020914-0000-0000-C000-000000000046
var IID_TablesOfContents = syscall.GUID{0x00020914, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TablesOfContents struct {
	ole.OleClient
}

func NewTablesOfContents(pDisp *win32.IDispatch, addRef bool, scoped bool) *TablesOfContents {
	p := &TablesOfContents{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TablesOfContentsFromVar(v ole.Variant) *TablesOfContents {
	return NewTablesOfContents(v.PdispValVal(), false, false)
}

func (this *TablesOfContents) IID() *syscall.GUID {
	return &IID_TablesOfContents
}

func (this *TablesOfContents) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TablesOfContents) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *TablesOfContents) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TablesOfContents) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *TablesOfContents) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *TablesOfContents) ForEach(action func(item *TableOfContents) bool) {
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
		pItem := (*TableOfContents)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *TablesOfContents) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *TablesOfContents) Format() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *TablesOfContents) SetFormat(rhs int32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *TablesOfContents) Item(index int32) *TableOfContents {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewTableOfContents(retVal.PdispValVal(), false, true)
}

var TablesOfContents_AddOld_OptArgs= []string{
	"UseHeadingStyles", "UpperHeadingLevel", "LowerHeadingLevel", "UseFields", 
	"TableID", "RightAlignPageNumbers", "IncludePageNumbers", "AddedStyles", 
}

func (this *TablesOfContents) AddOld(range_ *Range, optArgs ...interface{}) *TableOfContents {
	optArgs = ole.ProcessOptArgs(TablesOfContents_AddOld_OptArgs, optArgs)
	retVal := this.Call(0x00000064, []interface{}{range_}, optArgs...)
	return NewTableOfContents(retVal.PdispValVal(), false, true)
}

var TablesOfContents_MarkEntry_OptArgs= []string{
	"Entry", "EntryAutoText", "TableID", "Level", 
}

func (this *TablesOfContents) MarkEntry(range_ *Range, optArgs ...interface{}) *Field {
	optArgs = ole.ProcessOptArgs(TablesOfContents_MarkEntry_OptArgs, optArgs)
	retVal := this.Call(0x00000065, []interface{}{range_}, optArgs...)
	return NewField(retVal.PdispValVal(), false, true)
}

var TablesOfContents_Add2000_OptArgs= []string{
	"UseHeadingStyles", "UpperHeadingLevel", "LowerHeadingLevel", "UseFields", 
	"TableID", "RightAlignPageNumbers", "IncludePageNumbers", "AddedStyles", 
	"UseHyperlinks", "HidePageNumbersInWeb", 
}

func (this *TablesOfContents) Add2000(range_ *Range, optArgs ...interface{}) *TableOfContents {
	optArgs = ole.ProcessOptArgs(TablesOfContents_Add2000_OptArgs, optArgs)
	retVal := this.Call(0x00000066, []interface{}{range_}, optArgs...)
	return NewTableOfContents(retVal.PdispValVal(), false, true)
}

var TablesOfContents_Add_OptArgs= []string{
	"UseHeadingStyles", "UpperHeadingLevel", "LowerHeadingLevel", "UseFields", 
	"TableID", "RightAlignPageNumbers", "IncludePageNumbers", "AddedStyles", 
	"UseHyperlinks", "HidePageNumbersInWeb", "UseOutlineLevels", 
}

func (this *TablesOfContents) Add(range_ *Range, optArgs ...interface{}) *TableOfContents {
	optArgs = ole.ProcessOptArgs(TablesOfContents_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000067, []interface{}{range_}, optArgs...)
	return NewTableOfContents(retVal.PdispValVal(), false, true)
}

