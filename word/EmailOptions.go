package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209DB-0000-0000-C000-000000000046
var IID_EmailOptions = syscall.GUID{0x000209DB, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type EmailOptions struct {
	ole.OleClient
}

func NewEmailOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *EmailOptions {
	 if pDisp == nil {
		return nil;
	}
	p := &EmailOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EmailOptionsFromVar(v ole.Variant) *EmailOptions {
	return NewEmailOptions(v.IDispatch(), false, false)
}

func (this *EmailOptions) IID() *syscall.GUID {
	return &IID_EmailOptions
}

func (this *EmailOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *EmailOptions) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *EmailOptions) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *EmailOptions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *EmailOptions) UseThemeStyle() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetUseThemeStyle(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *EmailOptions) MarkCommentsWith() string {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EmailOptions) SetMarkCommentsWith(rhs string)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *EmailOptions) MarkComments() bool {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetMarkComments(rhs bool)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *EmailOptions) EmailSignature() *EmailSignature {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return NewEmailSignature(retVal.IDispatch(), false, true)
}

func (this *EmailOptions) ComposeStyle() *Style {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return NewStyle(retVal.IDispatch(), false, true)
}

func (this *EmailOptions) ReplyStyle() *Style {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return NewStyle(retVal.IDispatch(), false, true)
}

func (this *EmailOptions) ThemeName() string {
	retVal, _ := this.PropGet(0x00000072, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *EmailOptions) SetThemeName(rhs string)  {
	_ = this.PropPut(0x00000072, []interface{}{rhs})
}

func (this *EmailOptions) Dummy1() bool {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) Dummy2() bool {
	retVal, _ := this.PropGet(0x00000070, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) Dummy3()  {
	retVal, _ := this.Call(0x00000071, nil)
	_= retVal
}

func (this *EmailOptions) NewColorOnReply() bool {
	retVal, _ := this.PropGet(0x00000074, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetNewColorOnReply(rhs bool)  {
	_ = this.PropPut(0x00000074, []interface{}{rhs})
}

func (this *EmailOptions) PlainTextStyle() *Style {
	retVal, _ := this.PropGet(0x00000075, nil)
	return NewStyle(retVal.IDispatch(), false, true)
}

func (this *EmailOptions) UseThemeStyleOnReply() bool {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetUseThemeStyleOnReply(rhs bool)  {
	_ = this.PropPut(0x00000076, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeApplyHeadings() bool {
	retVal, _ := this.PropGet(0x00000104, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeApplyHeadings(rhs bool)  {
	_ = this.PropPut(0x00000104, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeApplyBorders() bool {
	retVal, _ := this.PropGet(0x00000105, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeApplyBorders(rhs bool)  {
	_ = this.PropPut(0x00000105, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeApplyBulletedLists() bool {
	retVal, _ := this.PropGet(0x00000106, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeApplyBulletedLists(rhs bool)  {
	_ = this.PropPut(0x00000106, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeApplyNumberedLists() bool {
	retVal, _ := this.PropGet(0x00000107, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeApplyNumberedLists(rhs bool)  {
	_ = this.PropPut(0x00000107, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeReplaceQuotes() bool {
	retVal, _ := this.PropGet(0x00000108, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeReplaceQuotes(rhs bool)  {
	_ = this.PropPut(0x00000108, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeReplaceSymbols() bool {
	retVal, _ := this.PropGet(0x00000109, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeReplaceSymbols(rhs bool)  {
	_ = this.PropPut(0x00000109, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeReplaceOrdinals() bool {
	retVal, _ := this.PropGet(0x0000010a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeReplaceOrdinals(rhs bool)  {
	_ = this.PropPut(0x0000010a, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeReplaceFractions() bool {
	retVal, _ := this.PropGet(0x0000010b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeReplaceFractions(rhs bool)  {
	_ = this.PropPut(0x0000010b, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeReplacePlainTextEmphasis() bool {
	retVal, _ := this.PropGet(0x0000010c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeReplacePlainTextEmphasis(rhs bool)  {
	_ = this.PropPut(0x0000010c, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeFormatListItemBeginning() bool {
	retVal, _ := this.PropGet(0x0000010d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeFormatListItemBeginning(rhs bool)  {
	_ = this.PropPut(0x0000010d, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeDefineStyles() bool {
	retVal, _ := this.PropGet(0x0000010e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeDefineStyles(rhs bool)  {
	_ = this.PropPut(0x0000010e, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeReplaceHyperlinks() bool {
	retVal, _ := this.PropGet(0x00000110, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeReplaceHyperlinks(rhs bool)  {
	_ = this.PropPut(0x00000110, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeApplyTables() bool {
	retVal, _ := this.PropGet(0x00000122, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeApplyTables(rhs bool)  {
	_ = this.PropPut(0x00000122, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeApplyFirstIndents() bool {
	retVal, _ := this.PropGet(0x00000129, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeApplyFirstIndents(rhs bool)  {
	_ = this.PropPut(0x00000129, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeApplyDates() bool {
	retVal, _ := this.PropGet(0x0000012a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeApplyDates(rhs bool)  {
	_ = this.PropPut(0x0000012a, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeApplyClosings() bool {
	retVal, _ := this.PropGet(0x0000012b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeApplyClosings(rhs bool)  {
	_ = this.PropPut(0x0000012b, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeMatchParentheses() bool {
	retVal, _ := this.PropGet(0x0000012c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeMatchParentheses(rhs bool)  {
	_ = this.PropPut(0x0000012c, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeReplaceFarEastDashes() bool {
	retVal, _ := this.PropGet(0x0000012d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeReplaceFarEastDashes(rhs bool)  {
	_ = this.PropPut(0x0000012d, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeDeleteAutoSpaces() bool {
	retVal, _ := this.PropGet(0x0000012e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeDeleteAutoSpaces(rhs bool)  {
	_ = this.PropPut(0x0000012e, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeInsertClosings() bool {
	retVal, _ := this.PropGet(0x0000012f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeInsertClosings(rhs bool)  {
	_ = this.PropPut(0x0000012f, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeAutoLetterWizard() bool {
	retVal, _ := this.PropGet(0x00000130, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeAutoLetterWizard(rhs bool)  {
	_ = this.PropPut(0x00000130, []interface{}{rhs})
}

func (this *EmailOptions) AutoFormatAsYouTypeInsertOvers() bool {
	retVal, _ := this.PropGet(0x00000131, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetAutoFormatAsYouTypeInsertOvers(rhs bool)  {
	_ = this.PropPut(0x00000131, []interface{}{rhs})
}

func (this *EmailOptions) RelyOnCSS() bool {
	retVal, _ := this.PropGet(0x00000132, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetRelyOnCSS(rhs bool)  {
	_ = this.PropPut(0x00000132, []interface{}{rhs})
}

func (this *EmailOptions) HTMLFidelity() int32 {
	retVal, _ := this.PropGet(0x00000133, nil)
	return retVal.LValVal()
}

func (this *EmailOptions) SetHTMLFidelity(rhs int32)  {
	_ = this.PropPut(0x00000133, []interface{}{rhs})
}

func (this *EmailOptions) EmbedSmartTag() bool {
	retVal, _ := this.PropGet(0x00000134, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetEmbedSmartTag(rhs bool)  {
	_ = this.PropPut(0x00000134, []interface{}{rhs})
}

func (this *EmailOptions) TabIndentKey() bool {
	retVal, _ := this.PropGet(0x00000135, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *EmailOptions) SetTabIndentKey(rhs bool)  {
	_ = this.PropPut(0x00000135, []interface{}{rhs})
}

