package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020912-0000-0000-C000-000000000046
var IID_TablesOfAuthorities = syscall.GUID{0x00020912, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TablesOfAuthorities struct {
	ole.OleClient
}

func NewTablesOfAuthorities(pDisp *win32.IDispatch, addRef bool, scoped bool) *TablesOfAuthorities {
	 if pDisp == nil {
		return nil;
	}
	p := &TablesOfAuthorities{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TablesOfAuthoritiesFromVar(v ole.Variant) *TablesOfAuthorities {
	return NewTablesOfAuthorities(v.IDispatch(), false, false)
}

func (this *TablesOfAuthorities) IID() *syscall.GUID {
	return &IID_TablesOfAuthorities
}

func (this *TablesOfAuthorities) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TablesOfAuthorities) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TablesOfAuthorities) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TablesOfAuthorities) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TablesOfAuthorities) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *TablesOfAuthorities) ForEach(action func(item *TableOfAuthorities) bool) {
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
		pItem := (*TableOfAuthorities)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *TablesOfAuthorities) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *TablesOfAuthorities) Format() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *TablesOfAuthorities) SetFormat(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *TablesOfAuthorities) Item(index int32) *TableOfAuthorities {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewTableOfAuthorities(retVal.IDispatch(), false, true)
}

var TablesOfAuthorities_Add_OptArgs= []string{
	"Category", "Bookmark", "Passim", "KeepEntryFormatting", 
	"Separator", "IncludeSequenceName", "EntrySeparator", "PageRangeSeparator", 
	"IncludeCategoryHeader", "PageNumberSeparator", 
}

func (this *TablesOfAuthorities) Add(range_ *Range, optArgs ...interface{}) *TableOfAuthorities {
	optArgs = ole.ProcessOptArgs(TablesOfAuthorities_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000064, []interface{}{range_}, optArgs...)
	return NewTableOfAuthorities(retVal.IDispatch(), false, true)
}

func (this *TablesOfAuthorities) NextCitation(shortCitation string)  {
	retVal, _ := this.Call(0x00000067, []interface{}{shortCitation})
	_= retVal
}

var TablesOfAuthorities_MarkCitation_OptArgs= []string{
	"LongCitation", "LongCitationAutoText", "Category", 
}

func (this *TablesOfAuthorities) MarkCitation(range_ *Range, shortCitation string, optArgs ...interface{}) *Field {
	optArgs = ole.ProcessOptArgs(TablesOfAuthorities_MarkCitation_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, []interface{}{range_, shortCitation}, optArgs...)
	return NewField(retVal.IDispatch(), false, true)
}

var TablesOfAuthorities_MarkAllCitations_OptArgs= []string{
	"LongCitation", "LongCitationAutoText", "Category", 
}

func (this *TablesOfAuthorities) MarkAllCitations(shortCitation string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(TablesOfAuthorities_MarkAllCitations_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, []interface{}{shortCitation}, optArgs...)
	_= retVal
}

