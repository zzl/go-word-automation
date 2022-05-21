package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209C0-0000-0000-C000-000000000046
var IID_ListFormat = syscall.GUID{0x000209C0, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListFormat struct {
	ole.OleClient
}

func NewListFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListFormat {
	p := &ListFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListFormatFromVar(v ole.Variant) *ListFormat {
	return NewListFormat(v.PdispValVal(), false, false)
}

func (this *ListFormat) IID() *syscall.GUID {
	return &IID_ListFormat
}

func (this *ListFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListFormat) ListLevelNumber() int32 {
	retVal := this.PropGet(0x00000044, nil)
	return retVal.LValVal()
}

func (this *ListFormat) SetListLevelNumber(rhs int32)  {
	retVal := this.PropPut(0x00000044, []interface{}{rhs})
	_= retVal
}

func (this *ListFormat) List() *List {
	retVal := this.PropGet(0x00000045, nil)
	return NewList(retVal.PdispValVal(), false, true)
}

func (this *ListFormat) ListTemplate() *ListTemplate {
	retVal := this.PropGet(0x00000046, nil)
	return NewListTemplate(retVal.PdispValVal(), false, true)
}

func (this *ListFormat) ListValue() int32 {
	retVal := this.PropGet(0x00000047, nil)
	return retVal.LValVal()
}

func (this *ListFormat) SingleList() bool {
	retVal := this.PropGet(0x00000048, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListFormat) SingleListTemplate() bool {
	retVal := this.PropGet(0x00000049, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListFormat) ListType() int32 {
	retVal := this.PropGet(0x0000004a, nil)
	return retVal.LValVal()
}

func (this *ListFormat) ListString() string {
	retVal := this.PropGet(0x0000004b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListFormat) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ListFormat) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ListFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ListFormat) CanContinuePreviousList(listTemplate *ListTemplate) int32 {
	retVal := this.Call(0x000000b8, []interface{}{listTemplate})
	return retVal.LValVal()
}

var ListFormat_RemoveNumbers_OptArgs= []string{
	"NumberType", 
}

func (this *ListFormat) RemoveNumbers(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListFormat_RemoveNumbers_OptArgs, optArgs)
	retVal := this.Call(0x000000b9, nil, optArgs...)
	_= retVal
}

var ListFormat_ConvertNumbersToText_OptArgs= []string{
	"NumberType", 
}

func (this *ListFormat) ConvertNumbersToText(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListFormat_ConvertNumbersToText_OptArgs, optArgs)
	retVal := this.Call(0x000000ba, nil, optArgs...)
	_= retVal
}

var ListFormat_CountNumberedItems_OptArgs= []string{
	"NumberType", "Level", 
}

func (this *ListFormat) CountNumberedItems(optArgs ...interface{}) int32 {
	optArgs = ole.ProcessOptArgs(ListFormat_CountNumberedItems_OptArgs, optArgs)
	retVal := this.Call(0x000000bb, nil, optArgs...)
	return retVal.LValVal()
}

func (this *ListFormat) ApplyBulletDefaultOld()  {
	retVal := this.Call(0x000000bc, nil)
	_= retVal
}

func (this *ListFormat) ApplyNumberDefaultOld()  {
	retVal := this.Call(0x000000bd, nil)
	_= retVal
}

func (this *ListFormat) ApplyOutlineNumberDefaultOld()  {
	retVal := this.Call(0x000000be, nil)
	_= retVal
}

var ListFormat_ApplyListTemplateOld_OptArgs= []string{
	"ContinuePreviousList", "ApplyTo", 
}

func (this *ListFormat) ApplyListTemplateOld(listTemplate *ListTemplate, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListFormat_ApplyListTemplateOld_OptArgs, optArgs)
	retVal := this.Call(0x000000bf, []interface{}{listTemplate}, optArgs...)
	_= retVal
}

func (this *ListFormat) ListOutdent()  {
	retVal := this.Call(0x000000d2, nil)
	_= retVal
}

func (this *ListFormat) ListIndent()  {
	retVal := this.Call(0x000000d3, nil)
	_= retVal
}

var ListFormat_ApplyBulletDefault_OptArgs= []string{
	"DefaultListBehavior", 
}

func (this *ListFormat) ApplyBulletDefault(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListFormat_ApplyBulletDefault_OptArgs, optArgs)
	retVal := this.Call(0x000000d4, nil, optArgs...)
	_= retVal
}

var ListFormat_ApplyNumberDefault_OptArgs= []string{
	"DefaultListBehavior", 
}

func (this *ListFormat) ApplyNumberDefault(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListFormat_ApplyNumberDefault_OptArgs, optArgs)
	retVal := this.Call(0x000000d5, nil, optArgs...)
	_= retVal
}

var ListFormat_ApplyOutlineNumberDefault_OptArgs= []string{
	"DefaultListBehavior", 
}

func (this *ListFormat) ApplyOutlineNumberDefault(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListFormat_ApplyOutlineNumberDefault_OptArgs, optArgs)
	retVal := this.Call(0x000000d6, nil, optArgs...)
	_= retVal
}

var ListFormat_ApplyListTemplate_OptArgs= []string{
	"ContinuePreviousList", "ApplyTo", "DefaultListBehavior", 
}

func (this *ListFormat) ApplyListTemplate(listTemplate *ListTemplate, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListFormat_ApplyListTemplate_OptArgs, optArgs)
	retVal := this.Call(0x000000d7, []interface{}{listTemplate}, optArgs...)
	_= retVal
}

func (this *ListFormat) ListPictureBullet() *InlineShape {
	retVal := this.PropGet(0x0000004c, nil)
	return NewInlineShape(retVal.PdispValVal(), false, true)
}

var ListFormat_ApplyListTemplateWithLevel_OptArgs= []string{
	"ContinuePreviousList", "ApplyTo", "DefaultListBehavior", "ApplyLevel", 
}

func (this *ListFormat) ApplyListTemplateWithLevel(listTemplate *ListTemplate, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ListFormat_ApplyListTemplateWithLevel_OptArgs, optArgs)
	retVal := this.Call(0x000000d8, []interface{}{listTemplate}, optArgs...)
	_= retVal
}

