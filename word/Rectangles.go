package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 7D0F7985-68D9-4D93-91CB-8109280E76CC
var IID_Rectangles = syscall.GUID{0x7D0F7985, 0x68D9, 0x4D93, 
	[8]byte{0x91, 0xCB, 0x81, 0x09, 0x28, 0x0E, 0x76, 0xCC}}

type Rectangles struct {
	ole.OleClient
}

func NewRectangles(pDisp *win32.IDispatch, addRef bool, scoped bool) *Rectangles {
	p := &Rectangles{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RectanglesFromVar(v ole.Variant) *Rectangles {
	return NewRectangles(v.PdispValVal(), false, false)
}

func (this *Rectangles) IID() *syscall.GUID {
	return &IID_Rectangles
}

func (this *Rectangles) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Rectangles) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Rectangles) ForEach(action func(item *Rectangle) bool) {
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
		pItem := (*Rectangle)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Rectangles) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Rectangles) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Rectangles) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Rectangles) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Rectangles) Item(index int32) *Rectangle {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewRectangle(retVal.PdispValVal(), false, true)
}

