package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// ADD4EDF3-2F33-4734-9CE6-D476097C5ADA
var IID_Rectangle = syscall.GUID{0xADD4EDF3, 0x2F33, 0x4734, 
	[8]byte{0x9C, 0xE6, 0xD4, 0x76, 0x09, 0x7C, 0x5A, 0xDA}}

type Rectangle struct {
	ole.OleClient
}

func NewRectangle(pDisp *win32.IDispatch, addRef bool, scoped bool) *Rectangle {
	p := &Rectangle{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RectangleFromVar(v ole.Variant) *Rectangle {
	return NewRectangle(v.PdispValVal(), false, false)
}

func (this *Rectangle) IID() *syscall.GUID {
	return &IID_Rectangle
}

func (this *Rectangle) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Rectangle) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Rectangle) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Rectangle) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Rectangle) RectangleType() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Rectangle) Left() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Rectangle) Top() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Rectangle) Width() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Rectangle) Height() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Rectangle) Range() *Range {
	retVal := this.PropGet(0x00000007, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Rectangle) Lines() *Lines {
	retVal := this.PropGet(0x00000008, nil)
	return NewLines(retVal.PdispValVal(), false, true)
}

