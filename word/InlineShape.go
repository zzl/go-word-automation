package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209A8-0000-0000-C000-000000000046
var IID_InlineShape = syscall.GUID{0x000209A8, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type InlineShape struct {
	ole.OleClient
}

func NewInlineShape(pDisp *win32.IDispatch, addRef bool, scoped bool) *InlineShape {
	 if pDisp == nil {
		return nil;
	}
	p := &InlineShape{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func InlineShapeFromVar(v ole.Variant) *InlineShape {
	return NewInlineShape(v.IDispatch(), false, false)
}

func (this *InlineShape) IID() *syscall.GUID {
	return &IID_InlineShape
}

func (this *InlineShape) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *InlineShape) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *InlineShape) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *InlineShape) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *InlineShape) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *InlineShape) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *InlineShape) Range() *Range {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *InlineShape) LinkFormat() *LinkFormat {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewLinkFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) Field() *Field {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewField(retVal.IDispatch(), false, true)
}

func (this *InlineShape) OLEFormat() *OLEFormat {
	retVal, _ := this.PropGet(0x00000005, nil)
	return NewOLEFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) Type() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *InlineShape) Hyperlink() *Hyperlink {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewHyperlink(retVal.IDispatch(), false, true)
}

func (this *InlineShape) Height() float32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.FltValVal()
}

func (this *InlineShape) SetHeight(rhs float32)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *InlineShape) Width() float32 {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.FltValVal()
}

func (this *InlineShape) SetWidth(rhs float32)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *InlineShape) ScaleHeight() float32 {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.FltValVal()
}

func (this *InlineShape) SetScaleHeight(rhs float32)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *InlineShape) ScaleWidth() float32 {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return retVal.FltValVal()
}

func (this *InlineShape) SetScaleWidth(rhs float32)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *InlineShape) LockAspectRatio() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *InlineShape) SetLockAspectRatio(rhs int32)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *InlineShape) Line() *LineFormat {
	retVal, _ := this.PropGet(0x00000070, nil)
	return NewLineFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) Fill() *FillFormat {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return NewFillFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) PictureFormat() *PictureFormat {
	retVal, _ := this.PropGet(0x00000076, nil)
	return NewPictureFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) SetPictureFormat(rhs *PictureFormat)  {
	_ = this.PropPut(0x00000076, []interface{}{rhs})
}

func (this *InlineShape) Activate()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

func (this *InlineShape) Reset()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *InlineShape) Delete()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *InlineShape) Select()  {
	retVal, _ := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *InlineShape) ConvertToShape() *Shape {
	retVal, _ := this.Call(0x00000068, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *InlineShape) HorizontalLineFormat() *HorizontalLineFormat {
	retVal, _ := this.PropGet(0x00000077, nil)
	return NewHorizontalLineFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) Script() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *InlineShape) OWSAnchor() int32 {
	retVal, _ := this.PropGet(0x00000082, nil)
	return retVal.LValVal()
}

func (this *InlineShape) TextEffect() *TextEffectFormat {
	retVal, _ := this.PropGet(0x00000078, nil)
	return NewTextEffectFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) SetTextEffect(rhs *TextEffectFormat)  {
	_ = this.PropPut(0x00000078, []interface{}{rhs})
}

func (this *InlineShape) AlternativeText() string {
	retVal, _ := this.PropGet(0x00000083, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *InlineShape) SetAlternativeText(rhs string)  {
	_ = this.PropPut(0x00000083, []interface{}{rhs})
}

func (this *InlineShape) IsPictureBullet() bool {
	retVal, _ := this.PropGet(0x00000084, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *InlineShape) GroupItems() *GroupShapes {
	retVal, _ := this.PropGet(0x00000085, nil)
	return NewGroupShapes(retVal.IDispatch(), false, true)
}

func (this *InlineShape) HasChart() int32 {
	retVal, _ := this.PropGet(0x00000094, nil)
	return retVal.LValVal()
}

func (this *InlineShape) Chart() *Chart {
	retVal, _ := this.PropGet(0x00000095, nil)
	return NewChart(retVal.IDispatch(), false, true)
}

func (this *InlineShape) SoftEdge() *SoftEdgeFormat {
	retVal, _ := this.PropGet(0x00000098, nil)
	return NewSoftEdgeFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) Glow() *GlowFormat {
	retVal, _ := this.PropGet(0x00000099, nil)
	return NewGlowFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) Reflection() *ReflectionFormat {
	retVal, _ := this.PropGet(0x0000009a, nil)
	return NewReflectionFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) Shadow() *ShadowFormat {
	retVal, _ := this.PropGet(0x0000044d, nil)
	return NewShadowFormat(retVal.IDispatch(), false, true)
}

func (this *InlineShape) HasSmartArt() int32 {
	retVal, _ := this.PropGet(0x0000009b, nil)
	return retVal.LValVal()
}

func (this *InlineShape) SmartArt() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000009c, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *InlineShape) Title() string {
	retVal, _ := this.PropGet(0x0000009e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *InlineShape) SetTitle(rhs string)  {
	_ = this.PropPut(0x0000009e, []interface{}{rhs})
}

func (this *InlineShape) AnchorID() int32 {
	retVal, _ := this.PropGet(0x000000cf, nil)
	return retVal.LValVal()
}

func (this *InlineShape) EditID() int32 {
	retVal, _ := this.PropGet(0x000000d0, nil)
	return retVal.LValVal()
}

