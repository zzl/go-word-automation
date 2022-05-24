package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020922-0000-0000-C000-000000000046
var IID_TablesOfFigures = syscall.GUID{0x00020922, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TablesOfFigures struct {
	ole.OleClient
}

func NewTablesOfFigures(pDisp *win32.IDispatch, addRef bool, scoped bool) *TablesOfFigures {
	 if pDisp == nil {
		return nil;
	}
	p := &TablesOfFigures{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TablesOfFiguresFromVar(v ole.Variant) *TablesOfFigures {
	return NewTablesOfFigures(v.IDispatch(), false, false)
}

func (this *TablesOfFigures) IID() *syscall.GUID {
	return &IID_TablesOfFigures
}

func (this *TablesOfFigures) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TablesOfFigures) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TablesOfFigures) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TablesOfFigures) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TablesOfFigures) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *TablesOfFigures) ForEach(action func(item *TableOfFigures) bool) {
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
		pItem := (*TableOfFigures)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *TablesOfFigures) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *TablesOfFigures) Format() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *TablesOfFigures) SetFormat(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *TablesOfFigures) Item(index int32) *TableOfFigures {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewTableOfFigures(retVal.IDispatch(), false, true)
}

var TablesOfFigures_AddOld_OptArgs= []string{
	"Caption", "IncludeLabel", "UseHeadingStyles", "UpperHeadingLevel", 
	"LowerHeadingLevel", "UseFields", "TableID", "RightAlignPageNumbers", 
	"IncludePageNumbers", "AddedStyles", 
}

func (this *TablesOfFigures) AddOld(range_ *Range, optArgs ...interface{}) *TableOfFigures {
	optArgs = ole.ProcessOptArgs(TablesOfFigures_AddOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000064, []interface{}{range_}, optArgs...)
	return NewTableOfFigures(retVal.IDispatch(), false, true)
}

var TablesOfFigures_MarkEntry_OptArgs= []string{
	"Entry", "EntryAutoText", "TableID", "Level", 
}

func (this *TablesOfFigures) MarkEntry(range_ *Range, optArgs ...interface{}) *Field {
	optArgs = ole.ProcessOptArgs(TablesOfFigures_MarkEntry_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, []interface{}{range_}, optArgs...)
	return NewField(retVal.IDispatch(), false, true)
}

var TablesOfFigures_Add_OptArgs= []string{
	"Caption", "IncludeLabel", "UseHeadingStyles", "UpperHeadingLevel", 
	"LowerHeadingLevel", "UseFields", "TableID", "RightAlignPageNumbers", 
	"IncludePageNumbers", "AddedStyles", "UseHyperlinks", "HidePageNumbersInWeb", 
}

func (this *TablesOfFigures) Add(range_ *Range, optArgs ...interface{}) *TableOfFigures {
	optArgs = ole.ProcessOptArgs(TablesOfFigures_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bc, []interface{}{range_}, optArgs...)
	return NewTableOfFigures(retVal.IDispatch(), false, true)
}

