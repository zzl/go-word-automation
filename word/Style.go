package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002092C-0000-0000-C000-000000000046
var IID_Style = syscall.GUID{0x0002092C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Style struct {
	ole.OleClient
}

func NewStyle(pDisp *win32.IDispatch, addRef bool, scoped bool) *Style {
	p := &Style{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func StyleFromVar(v ole.Variant) *Style {
	return NewStyle(v.PdispValVal(), false, false)
}

func (this *Style) IID() *syscall.GUID {
	return &IID_Style
}

func (this *Style) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Style) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Style) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Style) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Style) NameLocal() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Style) SetNameLocal(rhs string)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *Style) BaseStyle() ole.Variant {
	retVal := this.PropGet(0x00000001, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Style) SetBaseStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *Style) Description() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Style) Type() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Style) BuiltIn() bool {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) NextParagraphStyle() ole.Variant {
	retVal := this.PropGet(0x00000005, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Style) SetNextParagraphStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *Style) InUse() bool {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) Shading() *Shading {
	retVal := this.PropGet(0x00000007, nil)
	return NewShading(retVal.PdispValVal(), false, true)
}

func (this *Style) Borders() *Borders {
	retVal := this.PropGet(0x00000008, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *Style) SetBorders(rhs *Borders)  {
	retVal := this.PropPut(0x00000008, []interface{}{rhs})
	_= retVal
}

func (this *Style) ParagraphFormat() *ParagraphFormat {
	retVal := this.PropGet(0x00000009, nil)
	return NewParagraphFormat(retVal.PdispValVal(), false, true)
}

func (this *Style) SetParagraphFormat(rhs *ParagraphFormat)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *Style) Font() *Font {
	retVal := this.PropGet(0x0000000a, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *Style) SetFont(rhs *Font)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *Style) Frame() *Frame {
	retVal := this.PropGet(0x0000000b, nil)
	return NewFrame(retVal.PdispValVal(), false, true)
}

func (this *Style) LanguageID() int32 {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *Style) SetLanguageID(rhs int32)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *Style) AutomaticallyUpdate() bool {
	retVal := this.PropGet(0x0000000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetAutomaticallyUpdate(rhs bool)  {
	retVal := this.PropPut(0x0000000d, []interface{}{rhs})
	_= retVal
}

func (this *Style) ListTemplate() *ListTemplate {
	retVal := this.PropGet(0x0000000e, nil)
	return NewListTemplate(retVal.PdispValVal(), false, true)
}

func (this *Style) ListLevelNumber() int32 {
	retVal := this.PropGet(0x0000000f, nil)
	return retVal.LValVal()
}

func (this *Style) LanguageIDFarEast() int32 {
	retVal := this.PropGet(0x00000010, nil)
	return retVal.LValVal()
}

func (this *Style) SetLanguageIDFarEast(rhs int32)  {
	retVal := this.PropPut(0x00000010, []interface{}{rhs})
	_= retVal
}

func (this *Style) Hidden() bool {
	retVal := this.PropGet(0x00000011, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetHidden(rhs bool)  {
	retVal := this.PropPut(0x00000011, []interface{}{rhs})
	_= retVal
}

func (this *Style) Delete()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

var Style_LinkToListTemplate_OptArgs= []string{
	"ListLevelNumber", 
}

func (this *Style) LinkToListTemplate(listTemplate *ListTemplate, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Style_LinkToListTemplate_OptArgs, optArgs)
	retVal := this.Call(0x00000065, []interface{}{listTemplate}, optArgs...)
	_= retVal
}

func (this *Style) NoProofing() int32 {
	retVal := this.PropGet(0x00000012, nil)
	return retVal.LValVal()
}

func (this *Style) SetNoProofing(rhs int32)  {
	retVal := this.PropPut(0x00000012, []interface{}{rhs})
	_= retVal
}

func (this *Style) LinkStyle() ole.Variant {
	retVal := this.PropGet(0x00000068, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Style) SetLinkStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *Style) Visibility() bool {
	retVal := this.PropGet(0x00000013, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetVisibility(rhs bool)  {
	retVal := this.PropPut(0x00000013, []interface{}{rhs})
	_= retVal
}

func (this *Style) NoSpaceBetweenParagraphsOfSameStyle() bool {
	retVal := this.PropGet(0x00000014, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetNoSpaceBetweenParagraphsOfSameStyle(rhs bool)  {
	retVal := this.PropPut(0x00000014, []interface{}{rhs})
	_= retVal
}

func (this *Style) Table() *TableStyle {
	retVal := this.PropGet(0x00000015, nil)
	return NewTableStyle(retVal.PdispValVal(), false, true)
}

func (this *Style) Locked() bool {
	retVal := this.PropGet(0x00000016, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x00000016, []interface{}{rhs})
	_= retVal
}

func (this *Style) Priority() int32 {
	retVal := this.PropGet(0x00000017, nil)
	return retVal.LValVal()
}

func (this *Style) SetPriority(rhs int32)  {
	retVal := this.PropPut(0x00000017, []interface{}{rhs})
	_= retVal
}

func (this *Style) UnhideWhenUsed() bool {
	retVal := this.PropGet(0x00000018, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetUnhideWhenUsed(rhs bool)  {
	retVal := this.PropPut(0x00000018, []interface{}{rhs})
	_= retVal
}

func (this *Style) QuickStyle() bool {
	retVal := this.PropGet(0x00000019, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Style) SetQuickStyle(rhs bool)  {
	retVal := this.PropPut(0x00000019, []interface{}{rhs})
	_= retVal
}

func (this *Style) Linked() bool {
	retVal := this.PropGet(0x0000001a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

