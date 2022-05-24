package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020957-0000-0000-C000-000000000046
var IID_Paragraph = syscall.GUID{0x00020957, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Paragraph struct {
	ole.OleClient
}

func NewParagraph(pDisp *win32.IDispatch, addRef bool, scoped bool) *Paragraph {
	 if pDisp == nil {
		return nil;
	}
	p := &Paragraph{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ParagraphFromVar(v ole.Variant) *Paragraph {
	return NewParagraph(v.IDispatch(), false, false)
}

func (this *Paragraph) IID() *syscall.GUID {
	return &IID_Paragraph
}

func (this *Paragraph) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Paragraph) Range() *Range {
	retVal, _ := this.PropGet(0x00000000, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Paragraph) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Paragraph) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Paragraph) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Paragraph) Format() *ParagraphFormat {
	retVal, _ := this.PropGet(0x0000044e, nil)
	return NewParagraphFormat(retVal.IDispatch(), false, true)
}

func (this *Paragraph) SetFormat(rhs *ParagraphFormat)  {
	_ = this.PropPut(0x0000044e, []interface{}{rhs})
}

func (this *Paragraph) TabStops() *TabStops {
	retVal, _ := this.PropGet(0x0000044f, nil)
	return NewTabStops(retVal.IDispatch(), false, true)
}

func (this *Paragraph) SetTabStops(rhs *TabStops)  {
	_ = this.PropPut(0x0000044f, []interface{}{rhs})
}

func (this *Paragraph) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Paragraph) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Paragraph) DropCap() *DropCap {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return NewDropCap(retVal.IDispatch(), false, true)
}

func (this *Paragraph) Style() ole.Variant {
	retVal, _ := this.PropGet(0x00000064, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Paragraph) SetStyle(rhs *ole.Variant)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *Paragraph) Alignment() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetAlignment(rhs int32)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *Paragraph) KeepTogether() int32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetKeepTogether(rhs int32)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *Paragraph) KeepWithNext() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetKeepWithNext(rhs int32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Paragraph) PageBreakBefore() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetPageBreakBefore(rhs int32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *Paragraph) NoLineNumber() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetNoLineNumber(rhs int32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *Paragraph) RightIndent() float32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetRightIndent(rhs float32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *Paragraph) LeftIndent() float32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetLeftIndent(rhs float32)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *Paragraph) FirstLineIndent() float32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetFirstLineIndent(rhs float32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *Paragraph) LineSpacing() float32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetLineSpacing(rhs float32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *Paragraph) LineSpacingRule() int32 {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetLineSpacingRule(rhs int32)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *Paragraph) SpaceBefore() float32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetSpaceBefore(rhs float32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *Paragraph) SpaceAfter() float32 {
	retVal, _ := this.PropGet(0x00000070, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetSpaceAfter(rhs float32)  {
	_ = this.PropPut(0x00000070, []interface{}{rhs})
}

func (this *Paragraph) Hyphenation() int32 {
	retVal, _ := this.PropGet(0x00000071, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetHyphenation(rhs int32)  {
	_ = this.PropPut(0x00000071, []interface{}{rhs})
}

func (this *Paragraph) WidowControl() int32 {
	retVal, _ := this.PropGet(0x00000072, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetWidowControl(rhs int32)  {
	_ = this.PropPut(0x00000072, []interface{}{rhs})
}

func (this *Paragraph) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Paragraph) FarEastLineBreakControl() int32 {
	retVal, _ := this.PropGet(0x00000075, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetFarEastLineBreakControl(rhs int32)  {
	_ = this.PropPut(0x00000075, []interface{}{rhs})
}

func (this *Paragraph) WordWrap() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetWordWrap(rhs int32)  {
	_ = this.PropPut(0x00000076, []interface{}{rhs})
}

func (this *Paragraph) HangingPunctuation() int32 {
	retVal, _ := this.PropGet(0x00000077, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetHangingPunctuation(rhs int32)  {
	_ = this.PropPut(0x00000077, []interface{}{rhs})
}

func (this *Paragraph) HalfWidthPunctuationOnTopOfLine() int32 {
	retVal, _ := this.PropGet(0x00000078, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetHalfWidthPunctuationOnTopOfLine(rhs int32)  {
	_ = this.PropPut(0x00000078, []interface{}{rhs})
}

func (this *Paragraph) AddSpaceBetweenFarEastAndAlpha() int32 {
	retVal, _ := this.PropGet(0x00000079, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetAddSpaceBetweenFarEastAndAlpha(rhs int32)  {
	_ = this.PropPut(0x00000079, []interface{}{rhs})
}

func (this *Paragraph) AddSpaceBetweenFarEastAndDigit() int32 {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetAddSpaceBetweenFarEastAndDigit(rhs int32)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *Paragraph) BaseLineAlignment() int32 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetBaseLineAlignment(rhs int32)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Paragraph) AutoAdjustRightIndent() int32 {
	retVal, _ := this.PropGet(0x0000007c, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetAutoAdjustRightIndent(rhs int32)  {
	_ = this.PropPut(0x0000007c, []interface{}{rhs})
}

func (this *Paragraph) DisableLineHeightGrid() int32 {
	retVal, _ := this.PropGet(0x0000007d, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetDisableLineHeightGrid(rhs int32)  {
	_ = this.PropPut(0x0000007d, []interface{}{rhs})
}

func (this *Paragraph) OutlineLevel() int32 {
	retVal, _ := this.PropGet(0x000000ca, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetOutlineLevel(rhs int32)  {
	_ = this.PropPut(0x000000ca, []interface{}{rhs})
}

func (this *Paragraph) CloseUp()  {
	retVal, _ := this.Call(0x0000012d, nil)
	_= retVal
}

func (this *Paragraph) OpenUp()  {
	retVal, _ := this.Call(0x0000012e, nil)
	_= retVal
}

func (this *Paragraph) OpenOrCloseUp()  {
	retVal, _ := this.Call(0x0000012f, nil)
	_= retVal
}

func (this *Paragraph) TabHangingIndent(count int16)  {
	retVal, _ := this.Call(0x00000130, []interface{}{count})
	_= retVal
}

func (this *Paragraph) TabIndent(count int16)  {
	retVal, _ := this.Call(0x00000132, []interface{}{count})
	_= retVal
}

func (this *Paragraph) Reset()  {
	retVal, _ := this.Call(0x00000138, nil)
	_= retVal
}

func (this *Paragraph) Space1()  {
	retVal, _ := this.Call(0x00000139, nil)
	_= retVal
}

func (this *Paragraph) Space15()  {
	retVal, _ := this.Call(0x0000013a, nil)
	_= retVal
}

func (this *Paragraph) Space2()  {
	retVal, _ := this.Call(0x0000013b, nil)
	_= retVal
}

func (this *Paragraph) IndentCharWidth(count int16)  {
	retVal, _ := this.Call(0x00000140, []interface{}{count})
	_= retVal
}

func (this *Paragraph) IndentFirstLineCharWidth(count int16)  {
	retVal, _ := this.Call(0x00000142, []interface{}{count})
	_= retVal
}

var Paragraph_Next_OptArgs= []string{
	"Count", 
}

func (this *Paragraph) Next(optArgs ...interface{}) *Paragraph {
	optArgs = ole.ProcessOptArgs(Paragraph_Next_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000144, nil, optArgs...)
	return NewParagraph(retVal.IDispatch(), false, true)
}

var Paragraph_Previous_OptArgs= []string{
	"Count", 
}

func (this *Paragraph) Previous(optArgs ...interface{}) *Paragraph {
	optArgs = ole.ProcessOptArgs(Paragraph_Previous_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000145, nil, optArgs...)
	return NewParagraph(retVal.IDispatch(), false, true)
}

func (this *Paragraph) OutlinePromote()  {
	retVal, _ := this.Call(0x00000146, nil)
	_= retVal
}

func (this *Paragraph) OutlineDemote()  {
	retVal, _ := this.Call(0x00000147, nil)
	_= retVal
}

func (this *Paragraph) OutlineDemoteToBody()  {
	retVal, _ := this.Call(0x00000148, nil)
	_= retVal
}

func (this *Paragraph) Indent()  {
	retVal, _ := this.Call(0x0000014d, nil)
	_= retVal
}

func (this *Paragraph) Outdent()  {
	retVal, _ := this.Call(0x0000014e, nil)
	_= retVal
}

func (this *Paragraph) CharacterUnitRightIndent() float32 {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetCharacterUnitRightIndent(rhs float32)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *Paragraph) CharacterUnitLeftIndent() float32 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetCharacterUnitLeftIndent(rhs float32)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *Paragraph) CharacterUnitFirstLineIndent() float32 {
	retVal, _ := this.PropGet(0x00000080, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetCharacterUnitFirstLineIndent(rhs float32)  {
	_ = this.PropPut(0x00000080, []interface{}{rhs})
}

func (this *Paragraph) LineUnitBefore() float32 {
	retVal, _ := this.PropGet(0x00000081, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetLineUnitBefore(rhs float32)  {
	_ = this.PropPut(0x00000081, []interface{}{rhs})
}

func (this *Paragraph) LineUnitAfter() float32 {
	retVal, _ := this.PropGet(0x00000082, nil)
	return retVal.FltValVal()
}

func (this *Paragraph) SetLineUnitAfter(rhs float32)  {
	_ = this.PropPut(0x00000082, []interface{}{rhs})
}

func (this *Paragraph) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x000000cb, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x000000cb, []interface{}{rhs})
}

func (this *Paragraph) ID() string {
	retVal, _ := this.PropGet(0x000000cc, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Paragraph) SetID(rhs string)  {
	_ = this.PropPut(0x000000cc, []interface{}{rhs})
}

func (this *Paragraph) SpaceBeforeAuto() int32 {
	retVal, _ := this.PropGet(0x00000084, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetSpaceBeforeAuto(rhs int32)  {
	_ = this.PropPut(0x00000084, []interface{}{rhs})
}

func (this *Paragraph) SpaceAfterAuto() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetSpaceAfterAuto(rhs int32)  {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *Paragraph) IsStyleSeparator() bool {
	retVal, _ := this.PropGet(0x00000086, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Paragraph) SelectNumber()  {
	retVal, _ := this.Call(0x0000014f, nil)
	_= retVal
}

var Paragraph_ListAdvanceTo_OptArgs= []string{
	"Level1", "Level2", "Level3", "Level4", 
	"Level5", "Level6", "Level7", "Level8", "Level9", 
}

func (this *Paragraph) ListAdvanceTo(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Paragraph_ListAdvanceTo_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000150, nil, optArgs...)
	_= retVal
}

func (this *Paragraph) ResetAdvanceTo()  {
	retVal, _ := this.Call(0x00000151, nil)
	_= retVal
}

func (this *Paragraph) SeparateList()  {
	retVal, _ := this.Call(0x00000152, nil)
	_= retVal
}

func (this *Paragraph) JoinList()  {
	retVal, _ := this.Call(0x00000153, nil)
	_= retVal
}

func (this *Paragraph) MirrorIndents() int32 {
	retVal, _ := this.PropGet(0x00000087, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetMirrorIndents(rhs int32)  {
	_ = this.PropPut(0x00000087, []interface{}{rhs})
}

func (this *Paragraph) TextboxTightWrap() int32 {
	retVal, _ := this.PropGet(0x00000088, nil)
	return retVal.LValVal()
}

func (this *Paragraph) SetTextboxTightWrap(rhs int32)  {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *Paragraph) ListNumberOriginal(level int16) int16 {
	retVal, _ := this.PropGet(0x00000089, []interface{}{level})
	return retVal.IValVal()
}

func (this *Paragraph) ParaID() int32 {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return retVal.LValVal()
}

func (this *Paragraph) TextID() int32 {
	retVal, _ := this.PropGet(0x0000008c, nil)
	return retVal.LValVal()
}

