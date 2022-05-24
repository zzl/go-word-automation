package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209B6-0000-0000-C000-000000000046
var IID_GroupShapes = syscall.GUID{0x000209B6, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type GroupShapes struct {
	ole.OleClient
}

func NewGroupShapes(pDisp *win32.IDispatch, addRef bool, scoped bool) *GroupShapes {
	 if pDisp == nil {
		return nil;
	}
	p := &GroupShapes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func GroupShapesFromVar(v ole.Variant) *GroupShapes {
	return NewGroupShapes(v.IDispatch(), false, false)
}

func (this *GroupShapes) IID() *syscall.GUID {
	return &IID_GroupShapes
}

func (this *GroupShapes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *GroupShapes) Application() *Application {
	retVal, _ := this.PropGet(0x00001f40, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *GroupShapes) Creator() int32 {
	retVal, _ := this.PropGet(0x00001f41, nil)
	return retVal.LValVal()
}

func (this *GroupShapes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GroupShapes) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *GroupShapes) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *GroupShapes) ForEach(action func(item *Shape) bool) {
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

func (this *GroupShapes) Item(index *ole.Variant) *Shape {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *GroupShapes) Range(index *ole.Variant) *ShapeRange {
	retVal, _ := this.Call(0x0000000a, []interface{}{index})
	return NewShapeRange(retVal.IDispatch(), false, true)
}

