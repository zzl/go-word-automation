package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209B5-0000-0000-C000-000000000046
var IID_ShapeRange = syscall.GUID{0x000209B5, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ShapeRange struct {
	ole.OleClient
}

func NewShapeRange(pDisp *win32.IDispatch, addRef bool, scoped bool) *ShapeRange {
	 if pDisp == nil {
		return nil;
	}
	p := &ShapeRange{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShapeRangeFromVar(v ole.Variant) *ShapeRange {
	return NewShapeRange(v.IDispatch(), false, false)
}

func (this *ShapeRange) IID() *syscall.GUID {
	return &IID_ShapeRange
}

func (this *ShapeRange) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ShapeRange) Application() *Application {
	retVal, _ := this.PropGet(0x00001f40, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Creator() int32 {
	retVal, _ := this.PropGet(0x00001f41, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ShapeRange) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ShapeRange) ForEach(action func(item *Shape) bool) {
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
		pItem := (*Shape)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ShapeRange) Adjustments() *Adjustments {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewAdjustments(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) AutoShapeType() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetAutoShapeType(rhs int32)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *ShapeRange) Callout() *CalloutFormat {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewCalloutFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) ConnectionSiteCount() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Connector() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) ConnectorFormat() *ConnectorFormat {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return NewConnectorFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Fill() *FillFormat {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return NewFillFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) GroupItems() *GroupShapes {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return NewGroupShapes(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Height() float32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetHeight(rhs float32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *ShapeRange) HorizontalFlip() int32 {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Left() float32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetLeft(rhs float32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *ShapeRange) Line() *LineFormat {
	retVal, _ := this.PropGet(0x00000070, nil)
	return NewLineFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) LockAspectRatio() int32 {
	retVal, _ := this.PropGet(0x00000071, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetLockAspectRatio(rhs int32)  {
	_ = this.PropPut(0x00000071, []interface{}{rhs})
}

func (this *ShapeRange) Name() string {
	retVal, _ := this.PropGet(0x00000073, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ShapeRange) SetName(rhs string)  {
	_ = this.PropPut(0x00000073, []interface{}{rhs})
}

func (this *ShapeRange) Nodes() *ShapeNodes {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewShapeNodes(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Rotation() float32 {
	retVal, _ := this.PropGet(0x00000075, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetRotation(rhs float32)  {
	_ = this.PropPut(0x00000075, []interface{}{rhs})
}

func (this *ShapeRange) PictureFormat() *PictureFormat {
	retVal, _ := this.PropGet(0x00000076, nil)
	return NewPictureFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Shadow() *ShadowFormat {
	retVal, _ := this.PropGet(0x00000077, nil)
	return NewShadowFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) TextEffect() *TextEffectFormat {
	retVal, _ := this.PropGet(0x00000078, nil)
	return NewTextEffectFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) TextFrame() *TextFrame {
	retVal, _ := this.PropGet(0x00000079, nil)
	return NewTextFrame(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) ThreeD() *ThreeDFormat {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return NewThreeDFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Top() float32 {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetTop(rhs float32)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *ShapeRange) Type() int32 {
	retVal, _ := this.PropGet(0x0000007c, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) VerticalFlip() int32 {
	retVal, _ := this.PropGet(0x0000007d, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Vertices() ole.Variant {
	retVal, _ := this.PropGet(0x0000007e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ShapeRange) Visible() int32 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetVisible(rhs int32)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *ShapeRange) Width() float32 {
	retVal, _ := this.PropGet(0x00000080, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetWidth(rhs float32)  {
	_ = this.PropPut(0x00000080, []interface{}{rhs})
}

func (this *ShapeRange) ZOrderPosition() int32 {
	retVal, _ := this.PropGet(0x00000081, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Hyperlink() *Hyperlink {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return NewHyperlink(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) RelativeHorizontalPosition() int32 {
	retVal, _ := this.PropGet(0x0000012c, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetRelativeHorizontalPosition(rhs int32)  {
	_ = this.PropPut(0x0000012c, []interface{}{rhs})
}

func (this *ShapeRange) RelativeVerticalPosition() int32 {
	retVal, _ := this.PropGet(0x0000012d, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetRelativeVerticalPosition(rhs int32)  {
	_ = this.PropPut(0x0000012d, []interface{}{rhs})
}

func (this *ShapeRange) LockAnchor() int32 {
	retVal, _ := this.PropGet(0x0000012e, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetLockAnchor(rhs int32)  {
	_ = this.PropPut(0x0000012e, []interface{}{rhs})
}

func (this *ShapeRange) WrapFormat() *WrapFormat {
	retVal, _ := this.PropGet(0x0000012f, nil)
	return NewWrapFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Anchor() *Range {
	retVal, _ := this.PropGet(0x00000130, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Item(index *ole.Variant) *Shape {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Align(align int32, relativeTo int32)  {
	retVal, _ := this.Call(0x0000000a, []interface{}{align, relativeTo})
	_= retVal
}

func (this *ShapeRange) Apply()  {
	retVal, _ := this.Call(0x0000000b, nil)
	_= retVal
}

func (this *ShapeRange) Delete()  {
	retVal, _ := this.Call(0x0000000c, nil)
	_= retVal
}

func (this *ShapeRange) Distribute(distribute int32, relativeTo int32)  {
	retVal, _ := this.Call(0x0000000d, []interface{}{distribute, relativeTo})
	_= retVal
}

func (this *ShapeRange) Duplicate() *ShapeRange {
	retVal, _ := this.Call(0x0000000e, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Flip(flipCmd int32)  {
	retVal, _ := this.Call(0x0000000f, []interface{}{flipCmd})
	_= retVal
}

func (this *ShapeRange) IncrementLeft(increment float32)  {
	retVal, _ := this.Call(0x00000010, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) IncrementRotation(increment float32)  {
	retVal, _ := this.Call(0x00000011, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) IncrementTop(increment float32)  {
	retVal, _ := this.Call(0x00000012, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) Group() *Shape {
	retVal, _ := this.Call(0x00000013, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) PickUp()  {
	retVal, _ := this.Call(0x00000014, nil)
	_= retVal
}

func (this *ShapeRange) Regroup() *Shape {
	retVal, _ := this.Call(0x00000015, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) RerouteConnections()  {
	retVal, _ := this.Call(0x00000016, nil)
	_= retVal
}

var ShapeRange_ScaleHeight_OptArgs= []string{
	"Scale", 
}

func (this *ShapeRange) ScaleHeight(factor float32, relativeToOriginalSize int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ShapeRange_ScaleHeight_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000017, []interface{}{factor, relativeToOriginalSize}, optArgs...)
	_= retVal
}

var ShapeRange_ScaleWidth_OptArgs= []string{
	"Scale", 
}

func (this *ShapeRange) ScaleWidth(factor float32, relativeToOriginalSize int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ShapeRange_ScaleWidth_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000018, []interface{}{factor, relativeToOriginalSize}, optArgs...)
	_= retVal
}

var ShapeRange_Select_OptArgs= []string{
	"Replace", 
}

func (this *ShapeRange) Select(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ShapeRange_Select_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000019, nil, optArgs...)
	_= retVal
}

func (this *ShapeRange) SetShapesDefaultProperties()  {
	retVal, _ := this.Call(0x0000001a, nil)
	_= retVal
}

func (this *ShapeRange) Ungroup() *ShapeRange {
	retVal, _ := this.Call(0x0000001b, nil)
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) ZOrder(zorderCmd int32)  {
	retVal, _ := this.Call(0x0000001c, []interface{}{zorderCmd})
	_= retVal
}

func (this *ShapeRange) ConvertToFrame() *Frame {
	retVal, _ := this.Call(0x0000001d, nil)
	return NewFrame(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) ConvertToInlineShape() *InlineShape {
	retVal, _ := this.Call(0x0000001e, nil)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Activate()  {
	retVal, _ := this.Call(0x00000032, nil)
	_= retVal
}

func (this *ShapeRange) AlternativeText() string {
	retVal, _ := this.PropGet(0x00000083, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ShapeRange) SetAlternativeText(rhs string)  {
	_ = this.PropPut(0x00000083, []interface{}{rhs})
}

func (this *ShapeRange) HasDiagram() int32 {
	retVal, _ := this.PropGet(0x00000084, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) Diagram() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000085, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ShapeRange) HasDiagramNode() int32 {
	retVal, _ := this.PropGet(0x00000086, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) DiagramNode() *DiagramNode {
	retVal, _ := this.PropGet(0x00000087, nil)
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Child() int32 {
	retVal, _ := this.PropGet(0x00000088, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) ParentGroup() *Shape {
	retVal, _ := this.PropGet(0x00000089, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) CanvasItems() *CanvasShapes {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return NewCanvasShapes(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) ID() int32 {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) CanvasCropLeft(increment float32)  {
	retVal, _ := this.Call(0x0000008c, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) CanvasCropTop(increment float32)  {
	retVal, _ := this.Call(0x0000008d, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) CanvasCropRight(increment float32)  {
	retVal, _ := this.Call(0x0000008e, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) CanvasCropBottom(increment float32)  {
	retVal, _ := this.Call(0x0000008f, []interface{}{increment})
	_= retVal
}

func (this *ShapeRange) SetRTF(rhs string)  {
	_ = this.PropPut(0x00000090, []interface{}{rhs})
}

func (this *ShapeRange) LayoutInCell() int32 {
	retVal, _ := this.PropGet(0x00000091, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetLayoutInCell(rhs int32)  {
	_ = this.PropPut(0x00000091, []interface{}{rhs})
}

func (this *ShapeRange) LeftRelative() float32 {
	retVal, _ := this.PropGet(0x000000c8, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetLeftRelative(rhs float32)  {
	_ = this.PropPut(0x000000c8, []interface{}{rhs})
}

func (this *ShapeRange) TopRelative() float32 {
	retVal, _ := this.PropGet(0x000000c9, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetTopRelative(rhs float32)  {
	_ = this.PropPut(0x000000c9, []interface{}{rhs})
}

func (this *ShapeRange) WidthRelative() float32 {
	retVal, _ := this.PropGet(0x000000ca, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetWidthRelative(rhs float32)  {
	_ = this.PropPut(0x000000ca, []interface{}{rhs})
}

func (this *ShapeRange) HeightRelative() float32 {
	retVal, _ := this.PropGet(0x000000cb, nil)
	return retVal.FltValVal()
}

func (this *ShapeRange) SetHeightRelative(rhs float32)  {
	_ = this.PropPut(0x000000cb, []interface{}{rhs})
}

func (this *ShapeRange) RelativeHorizontalSize() int32 {
	retVal, _ := this.PropGet(0x000000cc, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetRelativeHorizontalSize(rhs int32)  {
	_ = this.PropPut(0x000000cc, []interface{}{rhs})
}

func (this *ShapeRange) RelativeVerticalSize() int32 {
	retVal, _ := this.PropGet(0x000000cd, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetRelativeVerticalSize(rhs int32)  {
	_ = this.PropPut(0x000000cd, []interface{}{rhs})
}

func (this *ShapeRange) SoftEdge() *SoftEdgeFormat {
	retVal, _ := this.PropGet(0x00000098, nil)
	return NewSoftEdgeFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Glow() *GlowFormat {
	retVal, _ := this.PropGet(0x00000099, nil)
	return NewGlowFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) Reflection() *ReflectionFormat {
	retVal, _ := this.PropGet(0x0000009a, nil)
	return NewReflectionFormat(retVal.IDispatch(), false, true)
}

func (this *ShapeRange) TextFrame2() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000009b, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ShapeRange) ShapeStyle() int32 {
	retVal, _ := this.PropGet(0x00000096, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetShapeStyle(rhs int32)  {
	_ = this.PropPut(0x00000096, []interface{}{rhs})
}

func (this *ShapeRange) BackgroundStyle() int32 {
	retVal, _ := this.PropGet(0x00000097, nil)
	return retVal.LValVal()
}

func (this *ShapeRange) SetBackgroundStyle(rhs int32)  {
	_ = this.PropPut(0x00000097, []interface{}{rhs})
}

func (this *ShapeRange) Title() string {
	retVal, _ := this.PropGet(0x000000ce, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ShapeRange) SetTitle(rhs string)  {
	_ = this.PropPut(0x000000ce, []interface{}{rhs})
}

