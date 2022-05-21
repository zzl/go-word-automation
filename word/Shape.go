package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209A0-0000-0000-C000-000000000046
var IID_Shape = syscall.GUID{0x000209A0, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Shape struct {
	ole.OleClient
}

func NewShape(pDisp *win32.IDispatch, addRef bool, scoped bool) *Shape {
	p := &Shape{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShapeFromVar(v ole.Variant) *Shape {
	return NewShape(v.PdispValVal(), false, false)
}

func (this *Shape) IID() *syscall.GUID {
	return &IID_Shape
}

func (this *Shape) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Shape) Application() *Application {
	retVal := this.PropGet(0x00001f40, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Shape) Creator() int32 {
	retVal := this.PropGet(0x00001f41, nil)
	return retVal.LValVal()
}

func (this *Shape) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) Adjustments() *Adjustments {
	retVal := this.PropGet(0x00000064, nil)
	return NewAdjustments(retVal.PdispValVal(), false, true)
}

func (this *Shape) AutoShapeType() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Shape) SetAutoShapeType(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Callout() *CalloutFormat {
	retVal := this.PropGet(0x00000067, nil)
	return NewCalloutFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) ConnectionSiteCount() int32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *Shape) Connector() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *Shape) ConnectorFormat() *ConnectorFormat {
	retVal := this.PropGet(0x0000006a, nil)
	return NewConnectorFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Fill() *FillFormat {
	retVal := this.PropGet(0x0000006b, nil)
	return NewFillFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) GroupItems() *GroupShapes {
	retVal := this.PropGet(0x0000006c, nil)
	return NewGroupShapes(retVal.PdispValVal(), false, true)
}

func (this *Shape) Height() float32 {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetHeight(rhs float32)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *Shape) HorizontalFlip() int32 {
	retVal := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *Shape) Left() float32 {
	retVal := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetLeft(rhs float32)  {
	retVal := this.PropPut(0x0000006f, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Line() *LineFormat {
	retVal := this.PropGet(0x00000070, nil)
	return NewLineFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) LockAspectRatio() int32 {
	retVal := this.PropGet(0x00000071, nil)
	return retVal.LValVal()
}

func (this *Shape) SetLockAspectRatio(rhs int32)  {
	retVal := this.PropPut(0x00000071, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Name() string {
	retVal := this.PropGet(0x00000073, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Shape) SetName(rhs string)  {
	retVal := this.PropPut(0x00000073, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Nodes() *ShapeNodes {
	retVal := this.PropGet(0x00000074, nil)
	return NewShapeNodes(retVal.PdispValVal(), false, true)
}

func (this *Shape) Rotation() float32 {
	retVal := this.PropGet(0x00000075, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetRotation(rhs float32)  {
	retVal := this.PropPut(0x00000075, []interface{}{rhs})
	_= retVal
}

func (this *Shape) PictureFormat() *PictureFormat {
	retVal := this.PropGet(0x00000076, nil)
	return NewPictureFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Shadow() *ShadowFormat {
	retVal := this.PropGet(0x00000077, nil)
	return NewShadowFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) TextEffect() *TextEffectFormat {
	retVal := this.PropGet(0x00000078, nil)
	return NewTextEffectFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) TextFrame() *TextFrame {
	retVal := this.PropGet(0x00000079, nil)
	return NewTextFrame(retVal.PdispValVal(), false, true)
}

func (this *Shape) ThreeD() *ThreeDFormat {
	retVal := this.PropGet(0x0000007a, nil)
	return NewThreeDFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Top() float32 {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetTop(rhs float32)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Type() int32 {
	retVal := this.PropGet(0x0000007c, nil)
	return retVal.LValVal()
}

func (this *Shape) VerticalFlip() int32 {
	retVal := this.PropGet(0x0000007d, nil)
	return retVal.LValVal()
}

func (this *Shape) Vertices() ole.Variant {
	retVal := this.PropGet(0x0000007e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Shape) Visible() int32 {
	retVal := this.PropGet(0x0000007f, nil)
	return retVal.LValVal()
}

func (this *Shape) SetVisible(rhs int32)  {
	retVal := this.PropPut(0x0000007f, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Width() float32 {
	retVal := this.PropGet(0x00000080, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetWidth(rhs float32)  {
	retVal := this.PropPut(0x00000080, []interface{}{rhs})
	_= retVal
}

func (this *Shape) ZOrderPosition() int32 {
	retVal := this.PropGet(0x00000081, nil)
	return retVal.LValVal()
}

func (this *Shape) Hyperlink() *Hyperlink {
	retVal := this.PropGet(0x000003e9, nil)
	return NewHyperlink(retVal.PdispValVal(), false, true)
}

func (this *Shape) RelativeHorizontalPosition() int32 {
	retVal := this.PropGet(0x0000012c, nil)
	return retVal.LValVal()
}

func (this *Shape) SetRelativeHorizontalPosition(rhs int32)  {
	retVal := this.PropPut(0x0000012c, []interface{}{rhs})
	_= retVal
}

func (this *Shape) RelativeVerticalPosition() int32 {
	retVal := this.PropGet(0x0000012d, nil)
	return retVal.LValVal()
}

func (this *Shape) SetRelativeVerticalPosition(rhs int32)  {
	retVal := this.PropPut(0x0000012d, []interface{}{rhs})
	_= retVal
}

func (this *Shape) LockAnchor() int32 {
	retVal := this.PropGet(0x0000012e, nil)
	return retVal.LValVal()
}

func (this *Shape) SetLockAnchor(rhs int32)  {
	retVal := this.PropPut(0x0000012e, []interface{}{rhs})
	_= retVal
}

func (this *Shape) WrapFormat() *WrapFormat {
	retVal := this.PropGet(0x0000012f, nil)
	return NewWrapFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) OLEFormat() *OLEFormat {
	retVal := this.PropGet(0x000001f4, nil)
	return NewOLEFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Anchor() *Range {
	retVal := this.PropGet(0x000001f5, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Shape) LinkFormat() *LinkFormat {
	retVal := this.PropGet(0x000001f6, nil)
	return NewLinkFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Apply()  {
	retVal := this.Call(0x0000000a, nil)
	_= retVal
}

func (this *Shape) Delete()  {
	retVal := this.Call(0x0000000b, nil)
	_= retVal
}

func (this *Shape) Duplicate() *Shape {
	retVal := this.Call(0x0000000c, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shape) Flip(flipCmd int32)  {
	retVal := this.Call(0x0000000d, []interface{}{flipCmd})
	_= retVal
}

func (this *Shape) IncrementLeft(increment float32)  {
	retVal := this.Call(0x0000000e, []interface{}{increment})
	_= retVal
}

func (this *Shape) IncrementRotation(increment float32)  {
	retVal := this.Call(0x0000000f, []interface{}{increment})
	_= retVal
}

func (this *Shape) IncrementTop(increment float32)  {
	retVal := this.Call(0x00000010, []interface{}{increment})
	_= retVal
}

func (this *Shape) PickUp()  {
	retVal := this.Call(0x00000011, nil)
	_= retVal
}

func (this *Shape) RerouteConnections()  {
	retVal := this.Call(0x00000012, nil)
	_= retVal
}

func (this *Shape) ScaleHeight(factor float32, relativeToOriginalSize int32, scale int32)  {
	retVal := this.Call(0x00000013, []interface{}{factor, relativeToOriginalSize, scale})
	_= retVal
}

func (this *Shape) ScaleWidth(factor float32, relativeToOriginalSize int32, scale int32)  {
	retVal := this.Call(0x00000014, []interface{}{factor, relativeToOriginalSize, scale})
	_= retVal
}

var Shape_Select_OptArgs= []string{
	"Replace", 
}

func (this *Shape) Select(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Shape_Select_OptArgs, optArgs)
	retVal := this.Call(0x00000015, nil, optArgs...)
	_= retVal
}

func (this *Shape) SetShapesDefaultProperties()  {
	retVal := this.Call(0x00000016, nil)
	_= retVal
}

func (this *Shape) Ungroup() *ShapeRange {
	retVal := this.Call(0x00000017, nil)
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *Shape) ZOrder(zorderCmd int32)  {
	retVal := this.Call(0x00000018, []interface{}{zorderCmd})
	_= retVal
}

func (this *Shape) ConvertToInlineShape() *InlineShape {
	retVal := this.Call(0x00000019, nil)
	return NewInlineShape(retVal.PdispValVal(), false, true)
}

func (this *Shape) ConvertToFrame() *Frame {
	retVal := this.Call(0x0000001d, nil)
	return NewFrame(retVal.PdispValVal(), false, true)
}

func (this *Shape) Activate()  {
	retVal := this.Call(0x00000032, nil)
	_= retVal
}

func (this *Shape) AlternativeText() string {
	retVal := this.PropGet(0x00000083, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Shape) SetAlternativeText(rhs string)  {
	retVal := this.PropPut(0x00000083, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Script() *ole.DispatchClass {
	retVal := this.PropGet(0x000001f7, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) HasDiagram() int32 {
	retVal := this.PropGet(0x00000084, nil)
	return retVal.LValVal()
}

func (this *Shape) Diagram() *ole.DispatchClass {
	retVal := this.PropGet(0x00000085, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) HasDiagramNode() int32 {
	retVal := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *Shape) DiagramNode() *DiagramNode {
	retVal := this.PropGet(0x00000087, nil)
	return NewDiagramNode(retVal.PdispValVal(), false, true)
}

func (this *Shape) Child() int32 {
	retVal := this.PropGet(0x00000088, nil)
	return retVal.LValVal()
}

func (this *Shape) ParentGroup() *Shape {
	retVal := this.PropGet(0x00000089, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *Shape) CanvasItems() *CanvasShapes {
	retVal := this.PropGet(0x0000008a, nil)
	return NewCanvasShapes(retVal.PdispValVal(), false, true)
}

func (this *Shape) ID() int32 {
	retVal := this.PropGet(0x0000008b, nil)
	return retVal.LValVal()
}

func (this *Shape) CanvasCropLeft(increment float32)  {
	retVal := this.Call(0x0000008c, []interface{}{increment})
	_= retVal
}

func (this *Shape) CanvasCropTop(increment float32)  {
	retVal := this.Call(0x0000008d, []interface{}{increment})
	_= retVal
}

func (this *Shape) CanvasCropRight(increment float32)  {
	retVal := this.Call(0x0000008e, []interface{}{increment})
	_= retVal
}

func (this *Shape) CanvasCropBottom(increment float32)  {
	retVal := this.Call(0x0000008f, []interface{}{increment})
	_= retVal
}

func (this *Shape) SetRTF(rhs string)  {
	retVal := this.PropPut(0x00000090, []interface{}{rhs})
	_= retVal
}

func (this *Shape) LayoutInCell() int32 {
	retVal := this.PropGet(0x00000091, nil)
	return retVal.LValVal()
}

func (this *Shape) SetLayoutInCell(rhs int32)  {
	retVal := this.PropPut(0x00000091, []interface{}{rhs})
	_= retVal
}

func (this *Shape) HasChart() int32 {
	retVal := this.PropGet(0x00000094, nil)
	return retVal.LValVal()
}

func (this *Shape) Chart() *Chart {
	retVal := this.PropGet(0x00000095, nil)
	return NewChart(retVal.PdispValVal(), false, true)
}

func (this *Shape) LeftRelative() float32 {
	retVal := this.PropGet(0x000000c8, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetLeftRelative(rhs float32)  {
	retVal := this.PropPut(0x000000c8, []interface{}{rhs})
	_= retVal
}

func (this *Shape) TopRelative() float32 {
	retVal := this.PropGet(0x000000c9, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetTopRelative(rhs float32)  {
	retVal := this.PropPut(0x000000c9, []interface{}{rhs})
	_= retVal
}

func (this *Shape) WidthRelative() float32 {
	retVal := this.PropGet(0x000000ca, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetWidthRelative(rhs float32)  {
	retVal := this.PropPut(0x000000ca, []interface{}{rhs})
	_= retVal
}

func (this *Shape) HeightRelative() float32 {
	retVal := this.PropGet(0x000000cb, nil)
	return retVal.FltValVal()
}

func (this *Shape) SetHeightRelative(rhs float32)  {
	retVal := this.PropPut(0x000000cb, []interface{}{rhs})
	_= retVal
}

func (this *Shape) RelativeHorizontalSize() int32 {
	retVal := this.PropGet(0x000000cc, nil)
	return retVal.LValVal()
}

func (this *Shape) SetRelativeHorizontalSize(rhs int32)  {
	retVal := this.PropPut(0x000000cc, []interface{}{rhs})
	_= retVal
}

func (this *Shape) RelativeVerticalSize() int32 {
	retVal := this.PropGet(0x000000cd, nil)
	return retVal.LValVal()
}

func (this *Shape) SetRelativeVerticalSize(rhs int32)  {
	retVal := this.PropPut(0x000000cd, []interface{}{rhs})
	_= retVal
}

func (this *Shape) SoftEdge() *SoftEdgeFormat {
	retVal := this.PropGet(0x00000098, nil)
	return NewSoftEdgeFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Glow() *GlowFormat {
	retVal := this.PropGet(0x00000099, nil)
	return NewGlowFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) Reflection() *ReflectionFormat {
	retVal := this.PropGet(0x0000009a, nil)
	return NewReflectionFormat(retVal.PdispValVal(), false, true)
}

func (this *Shape) TextFrame2() *ole.DispatchClass {
	retVal := this.PropGet(0x0000009b, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) HasSmartArt() int32 {
	retVal := this.PropGet(0x000000ce, nil)
	return retVal.LValVal()
}

func (this *Shape) SmartArt() *ole.DispatchClass {
	retVal := this.PropGet(0x0000009c, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Shape) ShapeStyle() int32 {
	retVal := this.PropGet(0x00000096, nil)
	return retVal.LValVal()
}

func (this *Shape) SetShapeStyle(rhs int32)  {
	retVal := this.PropPut(0x00000096, []interface{}{rhs})
	_= retVal
}

func (this *Shape) BackgroundStyle() int32 {
	retVal := this.PropGet(0x00000097, nil)
	return retVal.LValVal()
}

func (this *Shape) SetBackgroundStyle(rhs int32)  {
	retVal := this.PropPut(0x00000097, []interface{}{rhs})
	_= retVal
}

func (this *Shape) Title() string {
	retVal := this.PropGet(0x0000009e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Shape) SetTitle(rhs string)  {
	retVal := this.PropPut(0x0000009e, []interface{}{rhs})
	_= retVal
}

func (this *Shape) AnchorID() int32 {
	retVal := this.PropGet(0x000000cf, nil)
	return retVal.LValVal()
}

func (this *Shape) EditID() int32 {
	retVal := this.PropGet(0x000000d0, nil)
	return retVal.LValVal()
}

