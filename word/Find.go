package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209B0-0000-0000-C000-000000000046
var IID_Find = syscall.GUID{0x000209B0, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Find struct {
	ole.OleClient
}

func NewFind(pDisp *win32.IDispatch, addRef bool, scoped bool) *Find {
	 if pDisp == nil {
		return nil;
	}
	p := &Find{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FindFromVar(v ole.Variant) *Find {
	return NewFind(v.IDispatch(), false, false)
}

func (this *Find) IID() *syscall.GUID {
	return &IID_Find
}

func (this *Find) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Find) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Find) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Find) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Find) Forward() bool {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetForward(rhs bool)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *Find) Font() *Font {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *Find) SetFont(rhs *Font)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *Find) Found() bool {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) MatchAllWordForms() bool {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchAllWordForms(rhs bool)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *Find) MatchCase() bool {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchCase(rhs bool)  {
	_ = this.PropPut(0x0000000e, []interface{}{rhs})
}

func (this *Find) MatchWildcards() bool {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchWildcards(rhs bool)  {
	_ = this.PropPut(0x0000000f, []interface{}{rhs})
}

func (this *Find) MatchSoundsLike() bool {
	retVal, _ := this.PropGet(0x00000010, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchSoundsLike(rhs bool)  {
	_ = this.PropPut(0x00000010, []interface{}{rhs})
}

func (this *Find) MatchWholeWord() bool {
	retVal, _ := this.PropGet(0x00000011, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchWholeWord(rhs bool)  {
	_ = this.PropPut(0x00000011, []interface{}{rhs})
}

func (this *Find) MatchFuzzy() bool {
	retVal, _ := this.PropGet(0x00000028, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchFuzzy(rhs bool)  {
	_ = this.PropPut(0x00000028, []interface{}{rhs})
}

func (this *Find) MatchByte() bool {
	retVal, _ := this.PropGet(0x00000029, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchByte(rhs bool)  {
	_ = this.PropPut(0x00000029, []interface{}{rhs})
}

func (this *Find) ParagraphFormat() *ParagraphFormat {
	retVal, _ := this.PropGet(0x00000012, nil)
	return NewParagraphFormat(retVal.IDispatch(), false, true)
}

func (this *Find) SetParagraphFormat(rhs *ParagraphFormat)  {
	_ = this.PropPut(0x00000012, []interface{}{rhs})
}

func (this *Find) Style() ole.Variant {
	retVal, _ := this.PropGet(0x00000013, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Find) SetStyle(rhs *ole.Variant)  {
	_ = this.PropPut(0x00000013, []interface{}{rhs})
}

func (this *Find) Text() string {
	retVal, _ := this.PropGet(0x00000016, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Find) SetText(rhs string)  {
	_ = this.PropPut(0x00000016, []interface{}{rhs})
}

func (this *Find) LanguageID() int32 {
	retVal, _ := this.PropGet(0x00000017, nil)
	return retVal.LValVal()
}

func (this *Find) SetLanguageID(rhs int32)  {
	_ = this.PropPut(0x00000017, []interface{}{rhs})
}

func (this *Find) Highlight() int32 {
	retVal, _ := this.PropGet(0x00000018, nil)
	return retVal.LValVal()
}

func (this *Find) SetHighlight(rhs int32)  {
	_ = this.PropPut(0x00000018, []interface{}{rhs})
}

func (this *Find) Replacement() *Replacement {
	retVal, _ := this.PropGet(0x00000019, nil)
	return NewReplacement(retVal.IDispatch(), false, true)
}

func (this *Find) Frame() *Frame {
	retVal, _ := this.PropGet(0x0000001a, nil)
	return NewFrame(retVal.IDispatch(), false, true)
}

func (this *Find) Wrap() int32 {
	retVal, _ := this.PropGet(0x0000001b, nil)
	return retVal.LValVal()
}

func (this *Find) SetWrap(rhs int32)  {
	_ = this.PropPut(0x0000001b, []interface{}{rhs})
}

func (this *Find) Format() bool {
	retVal, _ := this.PropGet(0x0000001c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetFormat(rhs bool)  {
	_ = this.PropPut(0x0000001c, []interface{}{rhs})
}

func (this *Find) LanguageIDFarEast() int32 {
	retVal, _ := this.PropGet(0x0000001d, nil)
	return retVal.LValVal()
}

func (this *Find) SetLanguageIDFarEast(rhs int32)  {
	_ = this.PropPut(0x0000001d, []interface{}{rhs})
}

func (this *Find) LanguageIDOther() int32 {
	retVal, _ := this.PropGet(0x0000003c, nil)
	return retVal.LValVal()
}

func (this *Find) SetLanguageIDOther(rhs int32)  {
	_ = this.PropPut(0x0000003c, []interface{}{rhs})
}

func (this *Find) CorrectHangulEndings() bool {
	retVal, _ := this.PropGet(0x0000003d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetCorrectHangulEndings(rhs bool)  {
	_ = this.PropPut(0x0000003d, []interface{}{rhs})
}

var Find_ExecuteOld_OptArgs= []string{
	"FindText", "MatchCase", "MatchWholeWord", "MatchWildcards", 
	"MatchSoundsLike", "MatchAllWordForms", "Forward", "Wrap", 
	"Format", "ReplaceWith", "Replace", 
}

func (this *Find) ExecuteOld(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Find_ExecuteOld_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000001e, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) ClearFormatting()  {
	retVal, _ := this.Call(0x0000001f, nil)
	_= retVal
}

func (this *Find) SetAllFuzzyOptions()  {
	retVal, _ := this.Call(0x00000020, nil)
	_= retVal
}

func (this *Find) ClearAllFuzzyOptions()  {
	retVal, _ := this.Call(0x00000021, nil)
	_= retVal
}

var Find_Execute_OptArgs= []string{
	"FindText", "MatchCase", "MatchWholeWord", "MatchWildcards", 
	"MatchSoundsLike", "MatchAllWordForms", "Forward", "Wrap", 
	"Format", "ReplaceWith", "Replace", "MatchKashida", 
	"MatchDiacritics", "MatchAlefHamza", "MatchControl", 
}

func (this *Find) Execute(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Find_Execute_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bc, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) NoProofing() int32 {
	retVal, _ := this.PropGet(0x00000022, nil)
	return retVal.LValVal()
}

func (this *Find) SetNoProofing(rhs int32)  {
	_ = this.PropPut(0x00000022, []interface{}{rhs})
}

func (this *Find) MatchKashida() bool {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchKashida(rhs bool)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *Find) MatchDiacritics() bool {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchDiacritics(rhs bool)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *Find) MatchAlefHamza() bool {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchAlefHamza(rhs bool)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *Find) MatchControl() bool {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchControl(rhs bool)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Find) MatchPhrase() bool {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchPhrase(rhs bool)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *Find) MatchPrefix() bool {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchPrefix(rhs bool)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *Find) MatchSuffix() bool {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetMatchSuffix(rhs bool)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *Find) IgnoreSpace() bool {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetIgnoreSpace(rhs bool)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *Find) IgnorePunct() bool {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetIgnorePunct(rhs bool)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

var Find_HitHighlight_OptArgs= []string{
	"HighlightColor", "TextColor", "MatchCase", "MatchWholeWord", 
	"MatchPrefix", "MatchSuffix", "MatchPhrase", "MatchWildcards", 
	"MatchSoundsLike", "MatchAllWordForms", "MatchByte", "MatchFuzzy", 
	"MatchKashida", "MatchDiacritics", "MatchAlefHamza", "MatchControl", 
	"IgnoreSpace", "IgnorePunct", "HanjaPhoneticHangul", 
}

func (this *Find) HitHighlight(findText *ole.Variant, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Find_HitHighlight_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bd, []interface{}{findText}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) ClearHitHighlight() bool {
	retVal, _ := this.Call(0x000001be, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Find_Execute2007_OptArgs= []string{
	"FindText", "MatchCase", "MatchWholeWord", "MatchWildcards", 
	"MatchSoundsLike", "MatchAllWordForms", "Forward", "Wrap", 
	"Format", "ReplaceWith", "Replace", "MatchKashida", 
	"MatchDiacritics", "MatchAlefHamza", "MatchControl", "MatchPrefix", 
	"MatchSuffix", "MatchPhrase", "IgnoreSpace", "IgnorePunct", 
}

func (this *Find) Execute2007(optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(Find_Execute2007_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001bf, nil, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) HanjaPhoneticHangul() bool {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Find) SetHanjaPhoneticHangul(rhs bool)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

