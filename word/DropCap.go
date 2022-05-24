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
	 if pDisp == nil {
		return nil;
	}
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
	return NewDropCap(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DropCap) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *DropCap) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DropCap) Position() int32 {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *DropCap) SetPosition(rhs int32)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *DropCap) FontName() string {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropCap) SetFontName(rhs string)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *DropCap) LinesToDrop() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *DropCap) SetLinesToDrop(rhs int32)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *DropCap) DistanceFromText() float32 {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.FltValVal()
}

func (this *DropCap) SetDistanceFromText(rhs float32)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *DropCap) Clear()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

func (this *DropCap) Enable()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

