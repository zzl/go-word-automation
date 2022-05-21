package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002091F-0000-0000-C000-000000000046
var IID_MailMergeFields = syscall.GUID{0x0002091F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailMergeFields struct {
	ole.OleClient
}

func NewMailMergeFields(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailMergeFields {
	p := &MailMergeFields{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailMergeFieldsFromVar(v ole.Variant) *MailMergeFields {
	return NewMailMergeFields(v.PdispValVal(), false, false)
}

func (this *MailMergeFields) IID() *syscall.GUID {
	return &IID_MailMergeFields
}

func (this *MailMergeFields) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailMergeFields) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MailMergeFields) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailMergeFields) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MailMergeFields) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *MailMergeFields) ForEach(action func(item *MailMergeField) bool) {
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
		pItem := (*MailMergeField)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *MailMergeFields) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *MailMergeFields) Item(index int32) *MailMergeField {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

func (this *MailMergeFields) Add(range_ *Range, name string) *MailMergeField {
	retVal := this.Call(0x00000065, []interface{}{range_, name})
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

var MailMergeFields_AddAsk_OptArgs= []string{
	"Prompt", "DefaultAskText", "AskOnce", 
}

func (this *MailMergeFields) AddAsk(range_ *Range, name string, optArgs ...interface{}) *MailMergeField {
	optArgs = ole.ProcessOptArgs(MailMergeFields_AddAsk_OptArgs, optArgs)
	retVal := this.Call(0x00000066, []interface{}{range_, name}, optArgs...)
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

var MailMergeFields_AddFillIn_OptArgs= []string{
	"Prompt", "DefaultFillInText", "AskOnce", 
}

func (this *MailMergeFields) AddFillIn(range_ *Range, optArgs ...interface{}) *MailMergeField {
	optArgs = ole.ProcessOptArgs(MailMergeFields_AddFillIn_OptArgs, optArgs)
	retVal := this.Call(0x00000067, []interface{}{range_}, optArgs...)
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

var MailMergeFields_AddIf_OptArgs= []string{
	"CompareTo", "TrueAutoText", "TrueText", "FalseAutoText", "FalseText", 
}

func (this *MailMergeFields) AddIf(range_ *Range, mergeField string, comparison int32, optArgs ...interface{}) *MailMergeField {
	optArgs = ole.ProcessOptArgs(MailMergeFields_AddIf_OptArgs, optArgs)
	retVal := this.Call(0x00000068, []interface{}{range_, mergeField, comparison}, optArgs...)
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

func (this *MailMergeFields) AddMergeRec(range_ *Range) *MailMergeField {
	retVal := this.Call(0x00000069, []interface{}{range_})
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

func (this *MailMergeFields) AddMergeSeq(range_ *Range) *MailMergeField {
	retVal := this.Call(0x0000006a, []interface{}{range_})
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

func (this *MailMergeFields) AddNext(range_ *Range) *MailMergeField {
	retVal := this.Call(0x0000006b, []interface{}{range_})
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

var MailMergeFields_AddNextIf_OptArgs= []string{
	"CompareTo", 
}

func (this *MailMergeFields) AddNextIf(range_ *Range, mergeField string, comparison int32, optArgs ...interface{}) *MailMergeField {
	optArgs = ole.ProcessOptArgs(MailMergeFields_AddNextIf_OptArgs, optArgs)
	retVal := this.Call(0x0000006c, []interface{}{range_, mergeField, comparison}, optArgs...)
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

var MailMergeFields_AddSet_OptArgs= []string{
	"ValueText", "ValueAutoText", 
}

func (this *MailMergeFields) AddSet(range_ *Range, name string, optArgs ...interface{}) *MailMergeField {
	optArgs = ole.ProcessOptArgs(MailMergeFields_AddSet_OptArgs, optArgs)
	retVal := this.Call(0x0000006d, []interface{}{range_, name}, optArgs...)
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

var MailMergeFields_AddSkipIf_OptArgs= []string{
	"CompareTo", 
}

func (this *MailMergeFields) AddSkipIf(range_ *Range, mergeField string, comparison int32, optArgs ...interface{}) *MailMergeField {
	optArgs = ole.ProcessOptArgs(MailMergeFields_AddSkipIf_OptArgs, optArgs)
	retVal := this.Call(0x0000006e, []interface{}{range_, mergeField, comparison}, optArgs...)
	return NewMailMergeField(retVal.PdispValVal(), false, true)
}

