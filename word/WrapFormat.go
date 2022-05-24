package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209C3-0000-0000-C000-000000000046
var IID_WrapFormat = syscall.GUID{0x000209C3, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type WrapFormat struct {
	ole.OleClient
}

func NewWrapFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *WrapFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &WrapFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WrapFormatFromVar(v ole.Variant) *WrapFormat {
	return NewWrapFormat(v.IDispatch(), false, false)
}

func (this *WrapFormat) IID() *syscall.GUID {
	return &IID_WrapFormat
}

func (this *WrapFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *WrapFormat) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *WrapFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *WrapFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *WrapFormat) Type() int32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *WrapFormat) SetType(rhs int32)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *WrapFormat) Side() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *WrapFormat) SetSide(rhs int32)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *WrapFormat) DistanceTop() float32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *WrapFormat) SetDistanceTop(rhs float32)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *WrapFormat) DistanceBottom() float32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *WrapFormat) SetDistanceBottom(rhs float32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *WrapFormat) DistanceLeft() float32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.FltValVal()
}

func (this *WrapFormat) SetDistanceLeft(rhs float32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *WrapFormat) DistanceRight() float32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.FltValVal()
}

func (this *WrapFormat) SetDistanceRight(rhs float32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *WrapFormat) AllowOverlap() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *WrapFormat) SetAllowOverlap(rhs int32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

