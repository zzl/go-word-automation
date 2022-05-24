package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020952-0000-0000-C000-000000000046
var IID_Font_ = syscall.GUID{0x00020952, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Font_ struct {
	ole.OleClient
}

func NewFont_(pDisp *win32.IDispatch, addRef bool, scoped bool) *Font_ {
	 if pDisp == nil {
		return nil;
	}
	p := &Font_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func Font_FromVar(v ole.Variant) *Font_ {
	return NewFont_(v.IDispatch(), false, false)
}

func (this *Font_) IID() *syscall.GUID {
	return &IID_Font_
}

func (this *Font_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Font_) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Font_) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Font_) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Font_) Duplicate() *Font {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *Font_) Bold() int32 {
	retVal, _ := this.PropGet(0x00000082, nil)
	return retVal.LValVal()
}

func (this *Font_) SetBold(rhs int32)  {
	_ = this.PropPut(0x00000082, []interface{}{rhs})
}

func (this *Font_) Italic() int32 {
	retVal, _ := this.PropGet(0x00000083, nil)
	return retVal.LValVal()
}

func (this *Font_) SetItalic(rhs int32)  {
	_ = this.PropPut(0x00000083, []interface{}{rhs})
}

func (this *Font_) Hidden() int32 {
	retVal, _ := this.PropGet(0x00000084, nil)
	return retVal.LValVal()
}

func (this *Font_) SetHidden(rhs int32)  {
	_ = this.PropPut(0x00000084, []interface{}{rhs})
}

func (this *Font_) SmallCaps() int32 {
	retVal, _ := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *Font_) SetSmallCaps(rhs int32)  {
	_ = this.PropPut(0x00000085, []interface{}{rhs})
}

func (this *Font_) AllCaps() int32 {
	retVal, _ := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *Font_) SetAllCaps(rhs int32)  {
	_ = this.PropPut(0x00000086, []interface{}{rhs})
}

func (this *Font_) StrikeThrough() int32 {
	retVal, _ := this.PropGet(0x00000087, nil)
	return retVal.LValVal()
}

func (this *Font_) SetStrikeThrough(rhs int32)  {
	_ = this.PropPut(0x00000087, []interface{}{rhs})
}

func (this *Font_) DoubleStrikeThrough() int32 {
	retVal, _ := this.PropGet(0x00000088, nil)
	return retVal.LValVal()
}

func (this *Font_) SetDoubleStrikeThrough(rhs int32)  {
	_ = this.PropPut(0x00000088, []interface{}{rhs})
}

func (this *Font_) ColorIndex() int32 {
	retVal, _ := this.PropGet(0x00000089, nil)
	return retVal.LValVal()
}

func (this *Font_) SetColorIndex(rhs int32)  {
	_ = this.PropPut(0x00000089, []interface{}{rhs})
}

func (this *Font_) Subscript() int32 {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return retVal.LValVal()
}

func (this *Font_) SetSubscript(rhs int32)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *Font_) Superscript() int32 {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return retVal.LValVal()
}

func (this *Font_) SetSuperscript(rhs int32)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

func (this *Font_) Underline() int32 {
	retVal, _ := this.PropGet(0x0000008c, nil)
	return retVal.LValVal()
}

func (this *Font_) SetUnderline(rhs int32)  {
	_ = this.PropPut(0x0000008c, []interface{}{rhs})
}

func (this *Font_) Size() float32 {
	retVal, _ := this.PropGet(0x0000008d, nil)
	return retVal.FltValVal()
}

func (this *Font_) SetSize(rhs float32)  {
	_ = this.PropPut(0x0000008d, []interface{}{rhs})
}

func (this *Font_) Name() string {
	retVal, _ := this.PropGet(0x0000008e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Font_) SetName(rhs string)  {
	_ = this.PropPut(0x0000008e, []interface{}{rhs})
}

func (this *Font_) Position() int32 {
	retVal, _ := this.PropGet(0x0000008f, nil)
	return retVal.LValVal()
}

func (this *Font_) SetPosition(rhs int32)  {
	_ = this.PropPut(0x0000008f, []interface{}{rhs})
}

func (this *Font_) Spacing() float32 {
	retVal, _ := this.PropGet(0x00000090, nil)
	return retVal.FltValVal()
}

func (this *Font_) SetSpacing(rhs float32)  {
	_ = this.PropPut(0x00000090, []interface{}{rhs})
}

func (this *Font_) Scaling() int32 {
	retVal, _ := this.PropGet(0x00000091, nil)
	return retVal.LValVal()
}

func (this *Font_) SetScaling(rhs int32)  {
	_ = this.PropPut(0x00000091, []interface{}{rhs})
}

func (this *Font_) Shadow() int32 {
	retVal, _ := this.PropGet(0x00000092, nil)
	return retVal.LValVal()
}

func (this *Font_) SetShadow(rhs int32)  {
	_ = this.PropPut(0x00000092, []interface{}{rhs})
}

func (this *Font_) Outline() int32 {
	retVal, _ := this.PropGet(0x00000093, nil)
	return retVal.LValVal()
}

func (this *Font_) SetOutline(rhs int32)  {
	_ = this.PropPut(0x00000093, []interface{}{rhs})
}

func (this *Font_) Emboss() int32 {
	retVal, _ := this.PropGet(0x00000094, nil)
	return retVal.LValVal()
}

func (this *Font_) SetEmboss(rhs int32)  {
	_ = this.PropPut(0x00000094, []interface{}{rhs})
}

func (this *Font_) Kerning() float32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.FltValVal()
}

func (this *Font_) SetKerning(rhs float32)  {
	_ = this.PropPut(0x00000095, []interface{}{rhs})
}

func (this *Font_) Engrave() int32 {
	retVal, _ := this.PropGet(0x00000096, nil)
	return retVal.LValVal()
}

func (this *Font_) SetEngrave(rhs int32)  {
	_ = this.PropPut(0x00000096, []interface{}{rhs})
}

func (this *Font_) Animation() int32 {
	retVal, _ := this.PropGet(0x00000097, nil)
	return retVal.LValVal()
}

func (this *Font_) SetAnimation(rhs int32)  {
	_ = this.PropPut(0x00000097, []interface{}{rhs})
}

func (this *Font_) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Font_) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Font_) Shading() *Shading {
	retVal, _ := this.PropGet(0x00000099, nil)
	return NewShading(retVal.IDispatch(), false, true)
}

func (this *Font_) EmphasisMark() int32 {
	retVal, _ := this.PropGet(0x0000009a, nil)
	return retVal.LValVal()
}

func (this *Font_) SetEmphasisMark(rhs int32)  {
	_ = this.PropPut(0x0000009a, []interface{}{rhs})
}

func (this *Font_) DisableCharacterSpaceGrid() bool {
	retVal, _ := this.PropGet(0x0000009b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Font_) SetDisableCharacterSpaceGrid(rhs bool)  {
	_ = this.PropPut(0x0000009b, []interface{}{rhs})
}

func (this *Font_) NameFarEast() string {
	retVal, _ := this.PropGet(0x0000009c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Font_) SetNameFarEast(rhs string)  {
	_ = this.PropPut(0x0000009c, []interface{}{rhs})
}

func (this *Font_) NameAscii() string {
	retVal, _ := this.PropGet(0x0000009d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Font_) SetNameAscii(rhs string)  {
	_ = this.PropPut(0x0000009d, []interface{}{rhs})
}

func (this *Font_) NameOther() string {
	retVal, _ := this.PropGet(0x0000009e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Font_) SetNameOther(rhs string)  {
	_ = this.PropPut(0x0000009e, []interface{}{rhs})
}

func (this *Font_) Grow()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

func (this *Font_) Shrink()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *Font_) Reset()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *Font_) SetAsTemplateDefault()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

func (this *Font_) Color() int32 {
	retVal, _ := this.PropGet(0x0000009f, nil)
	return retVal.LValVal()
}

func (this *Font_) SetColor(rhs int32)  {
	_ = this.PropPut(0x0000009f, []interface{}{rhs})
}

func (this *Font_) BoldBi() int32 {
	retVal, _ := this.PropGet(0x000000a0, nil)
	return retVal.LValVal()
}

func (this *Font_) SetBoldBi(rhs int32)  {
	_ = this.PropPut(0x000000a0, []interface{}{rhs})
}

func (this *Font_) ItalicBi() int32 {
	retVal, _ := this.PropGet(0x000000a1, nil)
	return retVal.LValVal()
}

func (this *Font_) SetItalicBi(rhs int32)  {
	_ = this.PropPut(0x000000a1, []interface{}{rhs})
}

func (this *Font_) SizeBi() float32 {
	retVal, _ := this.PropGet(0x000000a2, nil)
	return retVal.FltValVal()
}

func (this *Font_) SetSizeBi(rhs float32)  {
	_ = this.PropPut(0x000000a2, []interface{}{rhs})
}

func (this *Font_) NameBi() string {
	retVal, _ := this.PropGet(0x000000a3, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Font_) SetNameBi(rhs string)  {
	_ = this.PropPut(0x000000a3, []interface{}{rhs})
}

func (this *Font_) ColorIndexBi() int32 {
	retVal, _ := this.PropGet(0x000000a4, nil)
	return retVal.LValVal()
}

func (this *Font_) SetColorIndexBi(rhs int32)  {
	_ = this.PropPut(0x000000a4, []interface{}{rhs})
}

func (this *Font_) DiacriticColor() int32 {
	retVal, _ := this.PropGet(0x000000a5, nil)
	return retVal.LValVal()
}

func (this *Font_) SetDiacriticColor(rhs int32)  {
	_ = this.PropPut(0x000000a5, []interface{}{rhs})
}

func (this *Font_) UnderlineColor() int32 {
	retVal, _ := this.PropGet(0x000000a6, nil)
	return retVal.LValVal()
}

func (this *Font_) SetUnderlineColor(rhs int32)  {
	_ = this.PropPut(0x000000a6, []interface{}{rhs})
}

func (this *Font_) Glow() *GlowFormat {
	retVal, _ := this.PropGet(0x000000a7, nil)
	return NewGlowFormat(retVal.IDispatch(), false, true)
}

func (this *Font_) SetGlow(rhs *GlowFormat)  {
	_ = this.PropPut(0x000000a7, []interface{}{rhs})
}

func (this *Font_) Reflection() *ReflectionFormat {
	retVal, _ := this.PropGet(0x000000a8, nil)
	return NewReflectionFormat(retVal.IDispatch(), false, true)
}

func (this *Font_) SetReflection(rhs *ReflectionFormat)  {
	_ = this.PropPut(0x000000a8, []interface{}{rhs})
}

func (this *Font_) TextShadow() *ShadowFormat {
	retVal, _ := this.PropGet(0x000000a9, nil)
	return NewShadowFormat(retVal.IDispatch(), false, true)
}

func (this *Font_) SetTextShadow(rhs *ShadowFormat)  {
	_ = this.PropPut(0x000000a9, []interface{}{rhs})
}

func (this *Font_) Fill() *FillFormat {
	retVal, _ := this.PropGet(0x000000aa, nil)
	return NewFillFormat(retVal.IDispatch(), false, true)
}

func (this *Font_) SetFill(rhs *FillFormat)  {
	_ = this.PropPut(0x000000aa, []interface{}{rhs})
}

func (this *Font_) Line() *LineFormat {
	retVal, _ := this.PropGet(0x000000ab, nil)
	return NewLineFormat(retVal.IDispatch(), false, true)
}

func (this *Font_) SetLine(rhs *LineFormat)  {
	_ = this.PropPut(0x000000ab, []interface{}{rhs})
}

func (this *Font_) ThreeD() *ThreeDFormat {
	retVal, _ := this.PropGet(0x000000ac, nil)
	return NewThreeDFormat(retVal.IDispatch(), false, true)
}

func (this *Font_) SetThreeD(rhs *ThreeDFormat)  {
	_ = this.PropPut(0x000000ac, []interface{}{rhs})
}

func (this *Font_) TextColor() *ColorFormat {
	retVal, _ := this.PropGet(0x000000ad, nil)
	return NewColorFormat(retVal.IDispatch(), false, true)
}

func (this *Font_) Ligatures() int32 {
	retVal, _ := this.PropGet(0x000000ae, nil)
	return retVal.LValVal()
}

func (this *Font_) SetLigatures(rhs int32)  {
	_ = this.PropPut(0x000000ae, []interface{}{rhs})
}

func (this *Font_) NumberForm() int32 {
	retVal, _ := this.PropGet(0x000000af, nil)
	return retVal.LValVal()
}

func (this *Font_) SetNumberForm(rhs int32)  {
	_ = this.PropPut(0x000000af, []interface{}{rhs})
}

func (this *Font_) NumberSpacing() int32 {
	retVal, _ := this.PropGet(0x000000b0, nil)
	return retVal.LValVal()
}

func (this *Font_) SetNumberSpacing(rhs int32)  {
	_ = this.PropPut(0x000000b0, []interface{}{rhs})
}

func (this *Font_) ContextualAlternates() int32 {
	retVal, _ := this.PropGet(0x000000b1, nil)
	return retVal.LValVal()
}

func (this *Font_) SetContextualAlternates(rhs int32)  {
	_ = this.PropPut(0x000000b1, []interface{}{rhs})
}

func (this *Font_) StylisticSet() int32 {
	retVal, _ := this.PropGet(0x000000b2, nil)
	return retVal.LValVal()
}

func (this *Font_) SetStylisticSet(rhs int32)  {
	_ = this.PropPut(0x000000b2, []interface{}{rhs})
}

