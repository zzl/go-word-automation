package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002099F-0000-0000-C000-000000000046
var IID_Shapes = syscall.GUID{0x0002099F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Shapes struct {
	ole.OleClient
}

func NewShapes(pDisp *win32.IDispatch, addRef bool, scoped bool) *Shapes {
	 if pDisp == nil {
		return nil;
	}
	p := &Shapes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShapesFromVar(v ole.Variant) *Shapes {
	return NewShapes(v.IDispatch(), false, false)
}

func (this *Shapes) IID() *syscall.GUID {
	return &IID_Shapes
}

func (this *Shapes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Shapes) Application() *Application {
	retVal, _ := this.PropGet(0x00001f40, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Shapes) Creator() int32 {
	retVal, _ := this.PropGet(0x00001f41, nil)
	return retVal.LValVal()
}

func (this *Shapes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Shapes) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Shapes) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Shapes) ForEach(action func(item *Shape) bool) {
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

func (this *Shapes) Item(index *ole.Variant) *Shape {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddCallout_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddCallout(type_ int32, left float32, top float32, width float32, height float32, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddCallout_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000a, []interface{}{type_, left, top, width, height}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *Shapes) AddConnector(type_ int32, beginX float32, beginY float32, endX float32, endY float32) *Shape {
	retVal, _ := this.Call(0x0000000b, []interface{}{type_, beginX, beginY, endX, endY})
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddCurve_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddCurve(safeArrayOfPoints *ole.Variant, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddCurve_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000c, []interface{}{safeArrayOfPoints}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddLabel_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddLabel(orientation int32, left float32, top float32, width float32, height float32, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddLabel_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000d, []interface{}{orientation, left, top, width, height}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddLine_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddLine(beginX float32, beginY float32, endX float32, endY float32, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddLine_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000e, []interface{}{beginX, beginY, endX, endY}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddPicture_OptArgs= []string{
	"LinkToFile", "SaveWithDocument", "Left", "Top", 
	"Width", "Height", "Anchor", 
}

func (this *Shapes) AddPicture(fileName string, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000f, []interface{}{fileName}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddPolyline_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddPolyline(safeArrayOfPoints *ole.Variant, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddPolyline_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000010, []interface{}{safeArrayOfPoints}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddShape_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddShape(type_ int32, left float32, top float32, width float32, height float32, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddShape_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000011, []interface{}{type_, left, top, width, height}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddTextEffect_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddTextEffect(presetTextEffect int32, text string, fontName string, fontSize float32, fontBold int32, fontItalic int32, left float32, top float32, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddTextEffect_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000012, []interface{}{presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddTextbox_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddTextbox(orientation int32, left float32, top float32, width float32, height float32, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddTextbox_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000013, []interface{}{orientation, left, top, width, height}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *Shapes) BuildFreeform(editingType int32, x1 float32, y1 float32) *FreeformBuilder {
	retVal, _ := this.Call(0x00000014, []interface{}{editingType, x1, y1})
	return NewFreeformBuilder(retVal.IDispatch(), false, true)
}

func (this *Shapes) Range(index *ole.Variant) *ShapeRange {
	retVal, _ := this.Call(0x00000015, []interface{}{index})
	return NewShapeRange(retVal.IDispatch(), false, true)
}

func (this *Shapes) SelectAll()  {
	retVal, _ := this.Call(0x00000016, nil)
	_= retVal
}

var Shapes_AddOLEObject_OptArgs= []string{
	"ClassType", "FileName", "LinkToFile", "DisplayAsIcon", 
	"IconFileName", "IconIndex", "IconLabel", "Left", 
	"Top", "Width", "Height", "Anchor", 
}

func (this *Shapes) AddOLEObject(optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddOLEObject_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000018, nil, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddOLEControl_OptArgs= []string{
	"ClassType", "Left", "Top", "Width", 
	"Height", "Anchor", 
}

func (this *Shapes) AddOLEControl(optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddOLEControl_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, nil, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddDiagram_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddDiagram(type_ int32, left float32, top float32, width float32, height float32, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddDiagram_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000017, []interface{}{type_, left, top, width, height}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddCanvas_OptArgs= []string{
	"Anchor", 
}

func (this *Shapes) AddCanvas(left float32, top float32, width float32, height float32, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddCanvas_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000019, []interface{}{left, top, width, height}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddChart_OptArgs= []string{
	"Type", "Left", "Top", "Width", 
	"Height", "Anchor", 
}

func (this *Shapes) AddChart(optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddChart_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000067, nil, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

var Shapes_AddSmartArt_OptArgs= []string{
	"Left", "Top", "Width", "Height", "Anchor", 
}

func (this *Shapes) AddSmartArt(layout *win32.IDispatch, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(Shapes_AddSmartArt_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000001c, []interface{}{layout}, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

