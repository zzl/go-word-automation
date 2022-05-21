package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020956-0000-0000-C000-000000000046
var IID_DropCap = syscall.GUID{0x00020956, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DropCap struct {
	ole.OleClient
}

func NewDropCap(pDisp *win32.IDispatch, addRef bool, scoped bool) *DropCap {
	p := &DropCap{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DropCapFromVar(v ole.Variant) *DropCap {
	return NewDropCap(v.PdispValVal(), false, false)
}

func (this *DropCap) IID() *syscall.GUID {
	return &IID_DropCap
}

func (this *DropCap) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DropCap) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *DropCap) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *DropCap) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DropCap) Position() int32 {
	retVal := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *DropCap) SetPosition(rhs int32)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *DropCap) FontName() string {
	retVal := this.PropGet(0x0000000b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropCap) SetFontName(rhs string)  {
	retVal := this.PropPut(0x0000000b, []interface{}{rhs})
	_= retVal
}

func (this *DropCap) LinesToDrop() int32 {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *DropCap) SetLinesToDrop(rhs int32)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *DropCap) DistanceFromText() float32 {
	retVal := this.PropGet(0x0000000d, nil)
	return retVal.FltValVal()
}

func (this *DropCap) SetDistanceFromText(rhs float32)  {
	retVal := this.PropPut(0x0000000d, []interface{}{rhs})
	_= retVal
}

func (this *DropCap) Clear()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

func (this *DropCap) Enable()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

