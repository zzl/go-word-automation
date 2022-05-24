package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// AE6CE2F5-B9D3-407D-85A8-0F10C63289A4
var IID_Line = syscall.GUID{0xAE6CE2F5, 0xB9D3, 0x407D, 
	[8]byte{0x85, 0xA8, 0x0F, 0x10, 0xC6, 0x32, 0x89, 0xA4}}

type Line struct {
	ole.OleClient
}

func NewLine(pDisp *win32.IDispatch, addRef bool, scoped bool) *Line {
	 if pDisp == nil {
		return nil;
	}
	p := &Line{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LineFromVar(v ole.Variant) *Line {
	return NewLine(v.IDispatch(), false, false)
}

func (this *Line) IID() *syscall.GUID {
	return &IID_Line
}

func (this *Line) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Line) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Line) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Line) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Line) LineType() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Line) Left() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Line) Top() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Line) Width() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Line) Height() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Line) Range() *Range {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Line) Rectangles() *Rectangles {
	retVal, _ := this.PropGet(0x00000008, nil)
	return NewRectangles(retVal.IDispatch(), false, true)
}

