package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// AE6D45E5-981E-4547-8752-674BB55420A5
var IID_Corners = syscall.GUID{0xAE6D45E5, 0x981E, 0x4547, 
	[8]byte{0x87, 0x52, 0x67, 0x4B, 0xB5, 0x54, 0x20, 0xA5}}

type Corners struct {
	ole.OleClient
}

func NewCorners(pDisp *win32.IDispatch, addRef bool, scoped bool) *Corners {
	 if pDisp == nil {
		return nil;
	}
	p := &Corners{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CornersFromVar(v ole.Variant) *Corners {
	return NewCorners(v.IDispatch(), false, false)
}

func (this *Corners) IID() *syscall.GUID {
	return &IID_Corners
}

func (this *Corners) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Corners) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Corners) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Corners) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Corners) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Corners) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

