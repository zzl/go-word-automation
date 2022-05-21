package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020958-0000-0000-C000-000000000046
var IID_Paragraphs = syscall.GUID{0x00020958, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Paragraphs struct {
	ole.OleClient
}

func NewParagraphs(pDisp *win32.IDispatch, addRef bool, scoped bool) *Paragraphs {
	p := &Paragraphs{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ParagraphsFromVar(v ole.Variant) *Paragraphs {
	return NewParagraphs(v.PdispValVal(), false, false)
}

func (this *Paragraphs) IID() *syscall.GUID {
	return &IID_Paragraphs
}

func (this *Paragraphs) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Paragraphs) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Paragraphs) ForEach(action func(item *Paragraph) bool) {
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
		pItem := (*Paragraph)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Paragraphs) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) First() *Paragraph {
	retVal := this.PropGet(0x00000003, nil)
	return NewParagraph(retVal.PdispValVal(), false, true)
}

func (this *Paragraphs) Last() *Paragraph {
	retVal := this.PropGet(0x00000004, nil)
	return NewParagraph(retVal.PdispValVal(), false, true)
}

func (this *Paragraphs) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Paragraphs) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Paragraphs) Format() *ParagraphFormat {
	retVal := this.PropGet(0x0000044e, nil)
	return NewParagraphFormat(retVal.PdispValVal(), false, true)
}

func (this *Paragraphs) SetFormat(rhs *ParagraphFormat)  {
	retVal := this.PropPut(0x0000044e, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) TabStops() *TabStops {
	retVal := this.PropGet(0x0000044f, nil)
	return NewTabStops(retVal.PdispValVal(), false, true)
}

func (this *Paragraphs) SetTabStops(rhs *TabStops)  {
	retVal := this.PropPut(0x0000044f, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) Borders() *Borders {
	retVal := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *Paragraphs) SetBorders(rhs *Borders)  {
	retVal := this.PropPut(0x0000044c, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) Style() ole.Variant {
	retVal := this.PropGet(0x00000064, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Paragraphs) SetStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) Alignment() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetAlignment(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) KeepTogether() int32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetKeepTogether(rhs int32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) KeepWithNext() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetKeepWithNext(rhs int32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) PageBreakBefore() int32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetPageBreakBefore(rhs int32)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) NoLineNumber() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetNoLineNumber(rhs int32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) RightIndent() float32 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetRightIndent(rhs float32)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) LeftIndent() float32 {
	retVal := this.PropGet(0x0000006b, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetLeftIndent(rhs float32)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) FirstLineIndent() float32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetFirstLineIndent(rhs float32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) LineSpacing() float32 {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetLineSpacing(rhs float32)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) LineSpacingRule() int32 {
	retVal := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetLineSpacingRule(rhs int32)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) SpaceBefore() float32 {
	retVal := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetSpaceBefore(rhs float32)  {
	retVal := this.PropPut(0x0000006f, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) SpaceAfter() float32 {
	retVal := this.PropGet(0x00000070, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetSpaceAfter(rhs float32)  {
	retVal := this.PropPut(0x00000070, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) Hyphenation() int32 {
	retVal := this.PropGet(0x00000071, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetHyphenation(rhs int32)  {
	retVal := this.PropPut(0x00000071, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) WidowControl() int32 {
	retVal := this.PropGet(0x00000072, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetWidowControl(rhs int32)  {
	retVal := this.PropPut(0x00000072, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) Shading() *Shading {
	retVal := this.PropGet(0x00000074, nil)
	return NewShading(retVal.PdispValVal(), false, true)
}

func (this *Paragraphs) FarEastLineBreakControl() int32 {
	retVal := this.PropGet(0x00000075, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetFarEastLineBreakControl(rhs int32)  {
	retVal := this.PropPut(0x00000075, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) WordWrap() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetWordWrap(rhs int32)  {
	retVal := this.PropPut(0x00000076, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) HangingPunctuation() int32 {
	retVal := this.PropGet(0x00000077, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetHangingPunctuation(rhs int32)  {
	retVal := this.PropPut(0x00000077, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) HalfWidthPunctuationOnTopOfLine() int32 {
	retVal := this.PropGet(0x00000078, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetHalfWidthPunctuationOnTopOfLine(rhs int32)  {
	retVal := this.PropPut(0x00000078, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) AddSpaceBetweenFarEastAndAlpha() int32 {
	retVal := this.PropGet(0x00000079, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetAddSpaceBetweenFarEastAndAlpha(rhs int32)  {
	retVal := this.PropPut(0x00000079, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) AddSpaceBetweenFarEastAndDigit() int32 {
	retVal := this.PropGet(0x0000007a, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetAddSpaceBetweenFarEastAndDigit(rhs int32)  {
	retVal := this.PropPut(0x0000007a, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) BaseLineAlignment() int32 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetBaseLineAlignment(rhs int32)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) AutoAdjustRightIndent() int32 {
	retVal := this.PropGet(0x0000007c, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetAutoAdjustRightIndent(rhs int32)  {
	retVal := this.PropPut(0x0000007c, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) DisableLineHeightGrid() int32 {
	retVal := this.PropGet(0x0000007d, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetDisableLineHeightGrid(rhs int32)  {
	retVal := this.PropPut(0x0000007d, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) OutlineLevel() int32 {
	retVal := this.PropGet(0x000000ca, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetOutlineLevel(rhs int32)  {
	retVal := this.PropPut(0x000000ca, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) Item(index int32) *Paragraph {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewParagraph(retVal.PdispValVal(), false, true)
}

var Paragraphs_Add_OptArgs= []string{
	"Range", 
}

func (this *Paragraphs) Add(optArgs ...interface{}) *Paragraph {
	optArgs = ole.ProcessOptArgs(Paragraphs_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000005, nil, optArgs...)
	return NewParagraph(retVal.PdispValVal(), false, true)
}

func (this *Paragraphs) CloseUp()  {
	retVal := this.Call(0x0000012d, nil)
	_= retVal
}

func (this *Paragraphs) OpenUp()  {
	retVal := this.Call(0x0000012e, nil)
	_= retVal
}

func (this *Paragraphs) OpenOrCloseUp()  {
	retVal := this.Call(0x0000012f, nil)
	_= retVal
}

func (this *Paragraphs) TabHangingIndent(count int16)  {
	retVal := this.Call(0x00000130, []interface{}{count})
	_= retVal
}

func (this *Paragraphs) TabIndent(count int16)  {
	retVal := this.Call(0x00000132, []interface{}{count})
	_= retVal
}

func (this *Paragraphs) Reset()  {
	retVal := this.Call(0x00000138, nil)
	_= retVal
}

func (this *Paragraphs) Space1()  {
	retVal := this.Call(0x00000139, nil)
	_= retVal
}

func (this *Paragraphs) Space15()  {
	retVal := this.Call(0x0000013a, nil)
	_= retVal
}

func (this *Paragraphs) Space2()  {
	retVal := this.Call(0x0000013b, nil)
	_= retVal
}

func (this *Paragraphs) IndentCharWidth(count int16)  {
	retVal := this.Call(0x00000140, []interface{}{count})
	_= retVal
}

func (this *Paragraphs) IndentFirstLineCharWidth(count int16)  {
	retVal := this.Call(0x00000142, []interface{}{count})
	_= retVal
}

func (this *Paragraphs) OutlinePromote()  {
	retVal := this.Call(0x00000144, nil)
	_= retVal
}

func (this *Paragraphs) OutlineDemote()  {
	retVal := this.Call(0x00000145, nil)
	_= retVal
}

func (this *Paragraphs) OutlineDemoteToBody()  {
	retVal := this.Call(0x00000146, nil)
	_= retVal
}

func (this *Paragraphs) Indent()  {
	retVal := this.Call(0x0000014d, nil)
	_= retVal
}

func (this *Paragraphs) Outdent()  {
	retVal := this.Call(0x0000014e, nil)
	_= retVal
}

func (this *Paragraphs) CharacterUnitRightIndent() float32 {
	retVal := this.PropGet(0x0000007e, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetCharacterUnitRightIndent(rhs float32)  {
	retVal := this.PropPut(0x0000007e, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) CharacterUnitLeftIndent() float32 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetCharacterUnitLeftIndent(rhs float32)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) CharacterUnitFirstLineIndent() float32 {
	retVal := this.PropGet(0x00000080, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetCharacterUnitFirstLineIndent(rhs float32)  {
	retVal := this.PropPut(0x00000080, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) LineUnitBefore() float32 {
	retVal := this.PropGet(0x00000081, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetLineUnitBefore(rhs float32)  {
	retVal := this.PropPut(0x00000081, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) LineUnitAfter() float32 {
	retVal := this.PropGet(0x00000082, nil)
	return retVal.FltValVal()
}

func (this *Paragraphs) SetLineUnitAfter(rhs float32)  {
	retVal := this.PropPut(0x00000082, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) ReadingOrder() int32 {
	retVal := this.PropGet(0x00000083, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x00000083, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) SpaceBeforeAuto() int32 {
	retVal := this.PropGet(0x00000084, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetSpaceBeforeAuto(rhs int32)  {
	retVal := this.PropPut(0x00000084, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) SpaceAfterAuto() int32 {
	retVal := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *Paragraphs) SetSpaceAfterAuto(rhs int32)  {
	retVal := this.PropPut(0x00000085, []interface{}{rhs})
	_= retVal
}

func (this *Paragraphs) IncreaseSpacing()  {
	retVal := this.Call(0x0000014f, nil)
	_= retVal
}

func (this *Paragraphs) DecreaseSpacing()  {
	retVal := this.Call(0x00000150, nil)
	_= retVal
}

