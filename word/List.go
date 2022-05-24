package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020992-0000-0000-C000-000000000046
var IID_List = syscall.GUID{0x00020992, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type List struct {
	ole.OleClient
}

func NewList(pDisp *win32.IDispatch, addRef bool, scoped bool) *List {
	 if pDisp == nil {
		return nil;
	}
	p := &List{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListFromVar(v ole.Variant) *List {
	return NewList(v.IDispatch(), false, false)
}

func (this *List) IID() *syscall.GUID {
	return &IID_List
}

func (this *List) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *List) Range() *Range {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *List) ListParagraphs() *ListParagraphs {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewListParagraphs(retVal.IDispatch(), false, true)
}

func (this *List) SingleListTemplate() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *List) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *List) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *List) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var List_ConvertNumbersToText_OptArgs= []string{
	"NumberType", 
}

func (this *List) ConvertNumbersToText(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(List_ConvertNumbersToText_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, nil, optArgs...)
	_= retVal
}

var List_RemoveNumbers_OptArgs= []string{
	"NumberType", 
}

func (this *List) RemoveNumbers(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(List_RemoveNumbers_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, nil, optArgs...)
	_= retVal
}

var List_CountNumberedItems_OptArgs= []string{
	"NumberType", "Level", 
}

func (this *List) CountNumberedItems(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(List_CountNumberedItems_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000067, nil, optArgs...)
	return retVal.LValVal()
}

var List_ApplyListTemplateOld_OptArgs= []string{
	"ContinuePreviousList", 
}

func (this *List) ApplyListTemplateOld(listTemplate *ListTemplate, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(List_ApplyListTemplateOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000068, []interface{}{listTemplate}, optArgs...)
	_= retVal
}

func (this *List) CanContinuePreviousList(listTemplate *ListTemplate) int32 {
	retVal, _ := this.Call(0x00000069, []interface{}{listTemplate})
	return retVal.LValVal()
}

var List_ApplyListTemplate_OptArgs= []string{
	"ContinuePreviousList", "DefaultListBehavior", 
}

func (this *List) ApplyListTemplate(listTemplate *ListTemplate, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(List_ApplyListTemplate_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006a, []interface{}{listTemplate}, optArgs...)
	_= retVal
}

func (this *List) StyleName() string {
	retVal, _ := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var List_ApplyListTemplateWithLevel_OptArgs= []string{
	"ContinuePreviousList", "DefaultListBehavior", "ApplyLevel", 
}

func (this *List) ApplyListTemplateWithLevel(listTemplate *ListTemplate, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(List_ApplyListTemplateWithLevel_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006b, []interface{}{listTemplate}, optArgs...)
	_= retVal
}

