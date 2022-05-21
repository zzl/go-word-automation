package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002097C-0000-0000-C000-000000000046
var IID_Indexes = syscall.GUID{0x0002097C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Indexes struct {
	ole.OleClient
}

func NewIndexes(pDisp *win32.IDispatch, addRef bool, scoped bool) *Indexes {
	p := &Indexes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func IndexesFromVar(v ole.Variant) *Indexes {
	return NewIndexes(v.PdispValVal(), false, false)
}

func (this *Indexes) IID() *syscall.GUID {
	return &IID_Indexes
}

func (this *Indexes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Indexes) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Indexes) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Indexes) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Indexes) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Indexes) ForEach(action func(item *Index) bool) {
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
		pItem := (*Index)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Indexes) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Indexes) Format() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Indexes) SetFormat(rhs int32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *Indexes) Item(index int32) *Index {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewIndex(retVal.PdispValVal(), false, true)
}

var Indexes_AddOld_OptArgs= []string{
	"HeadingSeparator", "RightAlignPageNumbers", "Type", "NumberOfColumns", "AccentedLetters", 
}

func (this *Indexes) AddOld(range_ *Range, optArgs ...interface{}) *Index {
	optArgs = ole.ProcessOptArgs(Indexes_AddOld_OptArgs, optArgs)
	retVal := this.Call(0x00000064, []interface{}{range_}, optArgs...)
	return NewIndex(retVal.PdispValVal(), false, true)
}

var Indexes_MarkEntry_OptArgs= []string{
	"Entry", "EntryAutoText", "CrossReference", "CrossReferenceAutoText", 
	"BookmarkName", "Bold", "Italic", "Reading", 
}

func (this *Indexes) MarkEntry(range_ *Range, optArgs ...interface{}) *Field {
	optArgs = ole.ProcessOptArgs(Indexes_MarkEntry_OptArgs, optArgs)
	retVal := this.Call(0x00000065, []interface{}{range_}, optArgs...)
	return NewField(retVal.PdispValVal(), false, true)
}

var Indexes_MarkAllEntries_OptArgs= []string{
	"Entry", "EntryAutoText", "CrossReference", "CrossReferenceAutoText", 
	"BookmarkName", "Bold", "Italic", 
}

func (this *Indexes) MarkAllEntries(range_ *Range, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Indexes_MarkAllEntries_OptArgs, optArgs)
	retVal := this.Call(0x00000066, []interface{}{range_}, optArgs...)
	_= retVal
}

func (this *Indexes) AutoMarkEntries(concordanceFileName string)  {
	retVal := this.Call(0x00000067, []interface{}{concordanceFileName})
	_= retVal
}

var Indexes_Add_OptArgs= []string{
	"HeadingSeparator", "RightAlignPageNumbers", "Type", "NumberOfColumns", 
	"AccentedLetters", "SortBy", "IndexLanguage", 
}

func (this *Indexes) Add(range_ *Range, optArgs ...interface{}) *Index {
	optArgs = ole.ProcessOptArgs(Indexes_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000068, []interface{}{range_}, optArgs...)
	return NewIndex(retVal.PdispValVal(), false, true)
}

