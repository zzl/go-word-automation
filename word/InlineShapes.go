package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209A9-0000-0000-C000-000000000046
var IID_InlineShapes = syscall.GUID{0x000209A9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type InlineShapes struct {
	ole.OleClient
}

func NewInlineShapes(pDisp *win32.IDispatch, addRef bool, scoped bool) *InlineShapes {
	 if pDisp == nil {
		return nil;
	}
	p := &InlineShapes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func InlineShapesFromVar(v ole.Variant) *InlineShapes {
	return NewInlineShapes(v.IDispatch(), false, false)
}

func (this *InlineShapes) IID() *syscall.GUID {
	return &IID_InlineShapes
}

func (this *InlineShapes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *InlineShapes) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *InlineShapes) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *InlineShapes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *InlineShapes) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *InlineShapes) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *InlineShapes) ForEach(action func(item *InlineShape) bool) {
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
		pItem := (*InlineShape)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *InlineShapes) Item(index int32) *InlineShape {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewInlineShape(retVal.IDispatch(), false, true)
}

var InlineShapes_AddPicture_OptArgs= []string{
	"LinkToFile", "SaveWithDocument", "Range", 
}

func (this *InlineShapes) AddPicture(fileName string, optArgs ...interface{}) *InlineShape {
	optArgs = ole.ProcessOptArgs(InlineShapes_AddPicture_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000064, []interface{}{fileName}, optArgs...)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

var InlineShapes_AddOLEObject_OptArgs= []string{
	"ClassType", "FileName", "LinkToFile", "DisplayAsIcon", 
	"IconFileName", "IconIndex", "IconLabel", "Range", 
}

func (this *InlineShapes) AddOLEObject(optArgs ...interface{}) *InlineShape {
	optArgs = ole.ProcessOptArgs(InlineShapes_AddOLEObject_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000018, nil, optArgs...)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

var InlineShapes_AddOLEControl_OptArgs= []string{
	"ClassType", "Range", 
}

func (this *InlineShapes) AddOLEControl(optArgs ...interface{}) *InlineShape {
	optArgs = ole.ProcessOptArgs(InlineShapes_AddOLEControl_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, nil, optArgs...)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

func (this *InlineShapes) New(range_ *Range) *InlineShape {
	retVal, _ := this.Call(0x000000c8, []interface{}{range_})
	return NewInlineShape(retVal.IDispatch(), false, true)
}

var InlineShapes_AddHorizontalLine_OptArgs= []string{
	"Range", 
}

func (this *InlineShapes) AddHorizontalLine(fileName string, optArgs ...interface{}) *InlineShape {
	optArgs = ole.ProcessOptArgs(InlineShapes_AddHorizontalLine_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000068, []interface{}{fileName}, optArgs...)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

var InlineShapes_AddHorizontalLineStandard_OptArgs= []string{
	"Range", 
}

func (this *InlineShapes) AddHorizontalLineStandard(optArgs ...interface{}) *InlineShape {
	optArgs = ole.ProcessOptArgs(InlineShapes_AddHorizontalLineStandard_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000069, nil, optArgs...)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

var InlineShapes_AddPictureBullet_OptArgs= []string{
	"Range", 
}

func (this *InlineShapes) AddPictureBullet(fileName string, optArgs ...interface{}) *InlineShape {
	optArgs = ole.ProcessOptArgs(InlineShapes_AddPictureBullet_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006a, []interface{}{fileName}, optArgs...)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

var InlineShapes_AddChart_OptArgs= []string{
	"Type", "Range", 
}

func (this *InlineShapes) AddChart(optArgs ...interface{}) *InlineShape {
	optArgs = ole.ProcessOptArgs(InlineShapes_AddChart_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006b, nil, optArgs...)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

var InlineShapes_AddSmartArt_OptArgs= []string{
	"Range", 
}

func (this *InlineShapes) AddSmartArt(layout *win32.IDispatch, optArgs ...interface{}) *InlineShape {
	optArgs = ole.ProcessOptArgs(InlineShapes_AddSmartArt_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006c, []interface{}{layout}, optArgs...)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

