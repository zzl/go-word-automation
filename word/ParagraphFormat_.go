package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020953-0000-0000-C000-000000000046
var IID_ParagraphFormat_ = syscall.GUID{0x00020953, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ParagraphFormat_ struct {
	ole.OleClient
}

func NewParagraphFormat_(pDisp *win32.IDispatch, addRef bool, scoped bool) *ParagraphFormat_ {
	p := &ParagraphFormat_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ParagraphFormat_FromVar(v ole.Variant) *ParagraphFormat_ {
	return NewParagraphFormat_(v.PdispValVal(), false, false)
}

func (this *ParagraphFormat_) IID() *syscall.GUID {
	return &IID_ParagraphFormat_
}

func (this *ParagraphFormat_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ParagraphFormat_) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ParagraphFormat_) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ParagraphFormat_) Duplicate() *ParagraphFormat {
	retVal := this.PropGet(0x0000000a, nil)
	return NewParagraphFormat(retVal.PdispValVal(), false, true)
}

func (this *ParagraphFormat_) Style() ole.Variant {
	retVal := this.PropGet(0x00000064, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ParagraphFormat_) SetStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) Alignment() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetAlignment(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) KeepTogether() int32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetKeepTogether(rhs int32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) KeepWithNext() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetKeepWithNext(rhs int32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) PageBreakBefore() int32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetPageBreakBefore(rhs int32)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) NoLineNumber() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetNoLineNumber(rhs int32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) RightIndent() float32 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetRightIndent(rhs float32)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) LeftIndent() float32 {
	retVal := this.PropGet(0x0000006b, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetLeftIndent(rhs float32)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) FirstLineIndent() float32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetFirstLineIndent(rhs float32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) LineSpacing() float32 {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetLineSpacing(rhs float32)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) LineSpacingRule() int32 {
	retVal := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetLineSpacingRule(rhs int32)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) SpaceBefore() float32 {
	retVal := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetSpaceBefore(rhs float32)  {
	retVal := this.PropPut(0x0000006f, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) SpaceAfter() float32 {
	retVal := this.PropGet(0x00000070, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetSpaceAfter(rhs float32)  {
	retVal := this.PropPut(0x00000070, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) Hyphenation() int32 {
	retVal := this.PropGet(0x00000071, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetHyphenation(rhs int32)  {
	retVal := this.PropPut(0x00000071, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) WidowControl() int32 {
	retVal := this.PropGet(0x00000072, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetWidowControl(rhs int32)  {
	retVal := this.PropPut(0x00000072, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) FarEastLineBreakControl() int32 {
	retVal := this.PropGet(0x00000075, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetFarEastLineBreakControl(rhs int32)  {
	retVal := this.PropPut(0x00000075, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) WordWrap() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetWordWrap(rhs int32)  {
	retVal := this.PropPut(0x00000076, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) HangingPunctuation() int32 {
	retVal := this.PropGet(0x00000077, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetHangingPunctuation(rhs int32)  {
	retVal := this.PropPut(0x00000077, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) HalfWidthPunctuationOnTopOfLine() int32 {
	retVal := this.PropGet(0x00000078, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetHalfWidthPunctuationOnTopOfLine(rhs int32)  {
	retVal := this.PropPut(0x00000078, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) AddSpaceBetweenFarEastAndAlpha() int32 {
	retVal := this.PropGet(0x00000079, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetAddSpaceBetweenFarEastAndAlpha(rhs int32)  {
	retVal := this.PropPut(0x00000079, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) AddSpaceBetweenFarEastAndDigit() int32 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetAddSpaceBetweenFarEastAndDigit(rhs int32)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) BaseLineAlignment() int32 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetBaseLineAlignment(rhs int32)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) AutoAdjustRightIndent() int32 {
	retVal := this.PropGet(0x0000007c, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetAutoAdjustRightIndent(rhs int32)  {
	retVal := this.PropPut(0x0000007c, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) DisableLineHeightGrid() int32 {
	retVal := this.PropGet(0x0000007d, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetDisableLineHeightGrid(rhs int32)  {
	retVal := this.PropPut(0x0000007d, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) TabStops() *TabStops {
	retVal := this.PropGet(0x0000044f, nil)
	return NewTabStops(retVal.PdispValVal(), false, true)
}

func (this *ParagraphFormat_) SetTabStops(rhs *TabStops)  {
	retVal := this.PropPut(0x0000044f, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) Borders() *Borders {
	retVal := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *ParagraphFormat_) SetBorders(rhs *Borders)  {
	retVal := this.PropPut(0x0000044c, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) Shading() *Shading {
	retVal := this.PropGet(0x0000044d, nil)
	return NewShading(retVal.PdispValVal(), false, true)
}

func (this *ParagraphFormat_) OutlineLevel() int32 {
	retVal := this.PropGet(0x000000ca, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetOutlineLevel(rhs int32)  {
	retVal := this.PropPut(0x000000ca, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) CloseUp()  {
	retVal := this.Call(0x0000012d, nil)
	_= retVal
}

func (this *ParagraphFormat_) OpenUp()  {
	retVal := this.Call(0x0000012e, nil)
	_= retVal
}

func (this *ParagraphFormat_) OpenOrCloseUp()  {
	retVal := this.Call(0x0000012f, nil)
	_= retVal
}

func (this *ParagraphFormat_) TabHangingIndent(count int16)  {
	retVal := this.Call(0x00000130, []interface{}{count})
	_= retVal
}

func (this *ParagraphFormat_) TabIndent(count int16)  {
	retVal := this.Call(0x00000132, []interface{}{count})
	_= retVal
}

func (this *ParagraphFormat_) Reset()  {
	retVal := this.Call(0x00000138, nil)
	_= retVal
}

func (this *ParagraphFormat_) Space1()  {
	retVal := this.Call(0x00000139, nil)
	_= retVal
}

func (this *ParagraphFormat_) Space15()  {
	retVal := this.Call(0x0000013a, nil)
	_= retVal
}

func (this *ParagraphFormat_) Space2()  {
	retVal := this.Call(0x0000013b, nil)
	_= retVal
}

func (this *ParagraphFormat_) IndentCharWidth(count int16)  {
	retVal := this.Call(0x00000140, []interface{}{count})
	_= retVal
}

func (this *ParagraphFormat_) IndentFirstLineCharWidth(count int16)  {
	retVal := this.Call(0x00000142, []interface{}{count})
	_= retVal
}

func (this *ParagraphFormat_) CharacterUnitRightIndent() float32 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetCharacterUnitRightIndent(rhs float32)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) CharacterUnitLeftIndent() float32 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetCharacterUnitLeftIndent(rhs float32)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) CharacterUnitFirstLineIndent() float32 {
	retVal := this.PropGet(0x00000080, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetCharacterUnitFirstLineIndent(rhs float32)  {
	retVal := this.PropPut(0x00000080, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) LineUnitBefore() float32 {
	retVal := this.PropGet(0x00000081, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetLineUnitBefore(rhs float32)  {
	retVal := this.PropPut(0x00000081, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) LineUnitAfter() float32 {
	retVal := this.PropGet(0x00000082, nil)
	return retVal.FltValVal()
}

func (this *ParagraphFormat_) SetLineUnitAfter(rhs float32)  {
	retVal := this.PropPut(0x00000082, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) ReadingOrder() int32 {
	retVal := this.PropGet(0x00000083, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x00000083, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) SpaceBeforeAuto() int32 {
	retVal := this.PropGet(0x00000084, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetSpaceBeforeAuto(rhs int32)  {
	retVal := this.PropPut(0x00000084, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) SpaceAfterAuto() int32 {
	retVal := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetSpaceAfterAuto(rhs int32)  {
	retVal := this.PropPut(0x00000085, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) MirrorIndents() int32 {
	retVal := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetMirrorIndents(rhs int32)  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *ParagraphFormat_) TextboxTightWrap() int32 {
	retVal := this.PropGet(0x00000087, nil)
	return retVal.LValVal()
}

func (this *ParagraphFormat_) SetTextboxTightWrap(rhs int32)  {
	retVal := this.PropPut(0x00000087, []interface{}{rhs})
	_= retVal
}

