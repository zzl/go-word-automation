package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 396F9073-F9FD-11D3-8EA0-0050049A1A01
var IID_CanvasShapes = syscall.GUID{0x396F9073, 0xF9FD, 0x11D3, 
	[8]byte{0x8E, 0xA0, 0x00, 0x50, 0x04, 0x9A, 0x1A, 0x01}}

type CanvasShapes struct {
	ole.OleClient
}

func NewCanvasShapes(pDisp *win32.IDispatch, addRef bool, scoped bool) *CanvasShapes {
	p := &CanvasShapes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CanvasShapesFromVar(v ole.Variant) *CanvasShapes {
	return NewCanvasShapes(v.PdispValVal(), false, false)
}

func (this *CanvasShapes) IID() *syscall.GUID {
	return &IID_CanvasShapes
}

func (this *CanvasShapes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CanvasShapes) Application() *Application {
	retVal := this.PropGet(0x00001f40, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) Creator() int32 {
	retVal := this.PropGet(0x00001f41, nil)
	return retVal.LValVal()
}

func (this *CanvasShapes) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CanvasShapes) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *CanvasShapes) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CanvasShapes) ForEach(action func(item *Shape) bool) {
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

func (this *CanvasShapes) Item(index *ole.Variant) *Shape {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) AddCallout(type_ int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x0000000a, []interface{}{type_, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) AddConnector(type_ int32, beginX float32, beginY float32, endX float32, endY float32) *Shape {
	retVal := this.Call(0x0000000b, []interface{}{type_, beginX, beginY, endX, endY})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) AddCurve(safeArrayOfPoints *ole.Variant) *Shape {
	retVal := this.Call(0x0000000c, []interface{}{safeArrayOfPoints})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) AddLabel(orientation int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x0000000d, []interface{}{orientation, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) AddLine(beginX float32, beginY float32, endX float32, endY float32) *Shape {
	retVal := this.Call(0x0000000e, []interface{}{beginX, beginY, endX, endY})
	return NewShape(retVal.PdispValVal(), false, true)
}

var CanvasShapes_AddPicture_OptArgs= []string{
	"LinkToFile", "SaveWithDocument", "Left", "Top", 
	"Width", "Height", 
}

func (this *CanvasShapes) AddPicture(fileName string, optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(CanvasShapes_AddPicture_OptArgs, optArgs)
	retVal := this.Call(0x0000000f, []interface{}{fileName}, optArgs...)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) AddPolyline(safeArrayOfPoints *ole.Variant) *Shape {
	retVal := this.Call(0x00000010, []interface{}{safeArrayOfPoints})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) AddShape(type_ int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x00000011, []interface{}{type_, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) AddTextEffect(presetTextEffect int32, text string, fontName string, fontSize float32, fontBold int32, fontItalic int32, left float32, top float32) *Shape {
	retVal := this.Call(0x00000012, []interface{}{presetTextEffect, text, fontName, fontSize, fontBold, fontItalic, left, top})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) AddTextbox(orientation int32, left float32, top float32, width float32, height float32) *Shape {
	retVal := this.Call(0x00000013, []interface{}{orientation, left, top, width, height})
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) BuildFreeform(editingType int32, x1 float32, y1 float32) *FreeformBuilder {
	retVal := this.Call(0x00000014, []interface{}{editingType, x1, y1})
	return NewFreeformBuilder(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) Range(index *ole.Variant) *ShapeRange {
	retVal := this.Call(0x00000015, []interface{}{index})
	return NewShapeRange(retVal.PdispValVal(), false, true)
}

func (this *CanvasShapes) SelectAll()  {
	retVal := this.Call(0x00000016, nil)
	_= retVal
}

