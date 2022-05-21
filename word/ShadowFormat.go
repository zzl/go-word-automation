package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209CC-0000-0000-C000-000000000046
var IID_ShadowFormat = syscall.GUID{0x000209CC, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ShadowFormat struct {
	ole.OleClient
}

func NewShadowFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ShadowFormat {
	p := &ShadowFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShadowFormatFromVar(v ole.Variant) *ShadowFormat {
	return NewShadowFormat(v.PdispValVal(), false, false)
}

func (this *ShadowFormat) IID() *syscall.GUID {
	return &IID_ShadowFormat
}

func (this *ShadowFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ShadowFormat) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ShadowFormat) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ShadowFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ShadowFormat) ForeColor() *ColorFormat {
	retVal := this.PropGet(0x00000064, nil)
	return NewColorFormat(retVal.PdispValVal(), false, true)
}

func (this *ShadowFormat) Obscured() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *ShadowFormat) SetObscured(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *ShadowFormat) OffsetX() float32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *ShadowFormat) SetOffsetX(rhs float32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *ShadowFormat) OffsetY() float32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *ShadowFormat) SetOffsetY(rhs float32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *ShadowFormat) Transparency() float32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.FltValVal()
}

func (this *ShadowFormat) SetTransparency(rhs float32)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *ShadowFormat) Type() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *ShadowFormat) SetType(rhs int32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *ShadowFormat) Visible() int32 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *ShadowFormat) SetVisible(rhs int32)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *ShadowFormat) IncrementOffsetX(increment float32)  {
	retVal := this.Call(0x0000000a, []interface{}{increment})
	_= retVal
}

func (this *ShadowFormat) IncrementOffsetY(increment float32)  {
	retVal := this.Call(0x0000000b, []interface{}{increment})
	_= retVal
}

func (this *ShadowFormat) Style() int32 {
	retVal := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *ShadowFormat) SetStyle(rhs int32)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

func (this *ShadowFormat) Blur() float32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.FltValVal()
}

func (this *ShadowFormat) SetBlur(rhs float32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *ShadowFormat) Size() float32 {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.FltValVal()
}

func (this *ShadowFormat) SetSize(rhs float32)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *ShadowFormat) RotateWithShape() int32 {
	retVal := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *ShadowFormat) SetRotateWithShape(rhs int32)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}
