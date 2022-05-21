package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 354AB591-A217-48B4-99E4-14F58F15667D
var IID_Axes = syscall.GUID{0x354AB591, 0xA217, 0x48B4, 
	[8]byte{0x99, 0xE4, 0x14, 0xF5, 0x8F, 0x15, 0x66, 0x7D}}

type Axes struct {
	ole.OleClient
}

func NewAxes(pDisp *win32.IDispatch, addRef bool, scoped bool) *Axes {
	p := &Axes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AxesFromVar(v ole.Variant) *Axes {
	return NewAxes(v.PdispValVal(), false, false)
}

func (this *Axes) IID() *syscall.GUID {
	return &IID_Axes
}

func (this *Axes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Axes) Count() int32 {
	retVal := this.PropGet(0x60020000, nil)
	return retVal.LValVal()
}

func (this *Axes) Item(type_ int32, axisGroup int32) *Axis {
	retVal := this.Call(0x00000000, []interface{}{type_, axisGroup})
	return NewAxis(retVal.PdispValVal(), false, true)
}

func (this *Axes) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Axes) ForEach(action func(item *Axis) bool) {
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
		pItem := (*Axis)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Axes) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Axes) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Axes) Default_(type_ int32, axisGroup int32) *Axis {
	retVal := this.Call(0x60020005, []interface{}{type_, axisGroup})
	return NewAxis(retVal.PdispValVal(), false, true)
}

func (this *Axes) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

