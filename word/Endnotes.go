package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020941-0000-0000-C000-000000000046
var IID_Endnotes = syscall.GUID{0x00020941, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Endnotes struct {
	ole.OleClient
}

func NewEndnotes(pDisp *win32.IDispatch, addRef bool, scoped bool) *Endnotes {
	p := &Endnotes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EndnotesFromVar(v ole.Variant) *Endnotes {
	return NewEndnotes(v.PdispValVal(), false, false)
}

func (this *Endnotes) IID() *syscall.GUID {
	return &IID_Endnotes
}

func (this *Endnotes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Endnotes) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Endnotes) ForEach(action func(item *Endnote) bool) {
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
		pItem := (*Endnote)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Endnotes) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Endnotes) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Endnotes) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Endnotes) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Endnotes) Location() int32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *Endnotes) SetLocation(rhs int32)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *Endnotes) NumberStyle() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Endnotes) SetNumberStyle(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *Endnotes) StartingNumber() int32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *Endnotes) SetStartingNumber(rhs int32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *Endnotes) NumberingRule() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *Endnotes) SetNumberingRule(rhs int32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *Endnotes) Separator() *Range {
	retVal := this.PropGet(0x00000068, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Endnotes) ContinuationSeparator() *Range {
	retVal := this.PropGet(0x00000069, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Endnotes) ContinuationNotice() *Range {
	retVal := this.PropGet(0x0000006a, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Endnotes) Item(index int32) *Endnote {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewEndnote(retVal.PdispValVal(), false, true)
}

var Endnotes_Add_OptArgs= []string{
	"Reference", "Text", 
}

func (this *Endnotes) Add(range_ *Range, optArgs ...interface{}) *Endnote {
	optArgs = ole.ProcessOptArgs(Endnotes_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000004, []interface{}{range_}, optArgs...)
	return NewEndnote(retVal.PdispValVal(), false, true)
}

func (this *Endnotes) Convert()  {
	retVal := this.Call(0x00000005, nil)
	_= retVal
}

func (this *Endnotes) SwapWithFootnotes()  {
	retVal := this.Call(0x00000006, nil)
	_= retVal
}

func (this *Endnotes) ResetSeparator()  {
	retVal := this.Call(0x00000007, nil)
	_= retVal
}

func (this *Endnotes) ResetContinuationSeparator()  {
	retVal := this.Call(0x00000008, nil)
	_= retVal
}

func (this *Endnotes) ResetContinuationNotice()  {
	retVal := this.Call(0x00000009, nil)
	_= retVal
}

