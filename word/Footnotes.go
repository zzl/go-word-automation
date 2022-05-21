package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020942-0000-0000-C000-000000000046
var IID_Footnotes = syscall.GUID{0x00020942, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Footnotes struct {
	ole.OleClient
}

func NewFootnotes(pDisp *win32.IDispatch, addRef bool, scoped bool) *Footnotes {
	p := &Footnotes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FootnotesFromVar(v ole.Variant) *Footnotes {
	return NewFootnotes(v.PdispValVal(), false, false)
}

func (this *Footnotes) IID() *syscall.GUID {
	return &IID_Footnotes
}

func (this *Footnotes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Footnotes) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Footnotes) ForEach(action func(item *Footnote) bool) {
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
		pItem := (*Footnote)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Footnotes) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Footnotes) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Footnotes) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Footnotes) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Footnotes) Location() int32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *Footnotes) SetLocation(rhs int32)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *Footnotes) NumberStyle() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Footnotes) SetNumberStyle(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *Footnotes) StartingNumber() int32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *Footnotes) SetStartingNumber(rhs int32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *Footnotes) NumberingRule() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *Footnotes) SetNumberingRule(rhs int32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *Footnotes) Separator() *Range {
	retVal := this.PropGet(0x00000068, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Footnotes) ContinuationSeparator() *Range {
	retVal := this.PropGet(0x00000069, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Footnotes) ContinuationNotice() *Range {
	retVal := this.PropGet(0x0000006a, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Footnotes) Item(index int32) *Footnote {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewFootnote(retVal.PdispValVal(), false, true)
}

var Footnotes_Add_OptArgs= []string{
	"Reference", "Text", 
}

func (this *Footnotes) Add(range_ *Range, optArgs ...interface{}) *Footnote {
	optArgs = ole.ProcessOptArgs(Footnotes_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000004, []interface{}{range_}, optArgs...)
	return NewFootnote(retVal.PdispValVal(), false, true)
}

func (this *Footnotes) Convert()  {
	retVal := this.Call(0x00000005, nil)
	_= retVal
}

func (this *Footnotes) SwapWithEndnotes()  {
	retVal := this.Call(0x00000006, nil)
	_= retVal
}

func (this *Footnotes) ResetSeparator()  {
	retVal := this.Call(0x00000007, nil)
	_= retVal
}

func (this *Footnotes) ResetContinuationSeparator()  {
	retVal := this.Call(0x00000008, nil)
	_= retVal
}

func (this *Footnotes) ResetContinuationNotice()  {
	retVal := this.Call(0x00000009, nil)
	_= retVal
}

