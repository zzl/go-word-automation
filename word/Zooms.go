package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209A7-0000-0000-C000-000000000046
var IID_Zooms = syscall.GUID{0x000209A7, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Zooms struct {
	ole.OleClient
}

func NewZooms(pDisp *win32.IDispatch, addRef bool, scoped bool) *Zooms {
	 if pDisp == nil {
		return nil;
	}
	p := &Zooms{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ZoomsFromVar(v ole.Variant) *Zooms {
	return NewZooms(v.IDispatch(), false, false)
}

func (this *Zooms) IID() *syscall.GUID {
	return &IID_Zooms
}

func (this *Zooms) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Zooms) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Zooms) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Zooms) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Zooms) Item(index int32) *Zoom {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewZoom(retVal.IDispatch(), false, true)
}

