package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// C1A870A0-850E-4D38-98A7-741CB8C3BCA4
var IID_Points = syscall.GUID{0xC1A870A0, 0x850E, 0x4D38, 
	[8]byte{0x98, 0xA7, 0x74, 0x1C, 0xB8, 0xC3, 0xBC, 0xA4}}

type Points struct {
	ole.OleClient
}

func NewPoints(pDisp *win32.IDispatch, addRef bool, scoped bool) *Points {
	 if pDisp == nil {
		return nil;
	}
	p := &Points{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PointsFromVar(v ole.Variant) *Points {
	return NewPoints(v.IDispatch(), false, false)
}

func (this *Points) IID() *syscall.GUID {
	return &IID_Points
}

func (this *Points) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Points) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Points) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *Points) Item(index int32) *Point {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewPoint(retVal.IDispatch(), false, true)
}

func (this *Points) NewEnum_() *com.UnknownClass {
	retVal, _ := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Points) ForEach(action func(item *Point) bool) {
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
		pItem := (*Point)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Points) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Points) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Points) Default_(index int32) *Point {
	retVal, _ := this.Call(0x60020006, []interface{}{index})
	return NewPoint(retVal.IDispatch(), false, true)
}

