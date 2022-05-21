package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209CB-0000-0000-C000-000000000046
var IID_PictureFormat = syscall.GUID{0x000209CB, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PictureFormat struct {
	ole.OleClient
}

func NewPictureFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *PictureFormat {
	p := &PictureFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PictureFormatFromVar(v ole.Variant) *PictureFormat {
	return NewPictureFormat(v.PdispValVal(), false, false)
}

func (this *PictureFormat) IID() *syscall.GUID {
	return &IID_PictureFormat
}

func (this *PictureFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PictureFormat) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *PictureFormat) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *PictureFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PictureFormat) Brightness() float32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.FltValVal()
}

func (this *PictureFormat) SetBrightness(rhs float32)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *PictureFormat) ColorType() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *PictureFormat) SetColorType(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *PictureFormat) Contrast() float32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *PictureFormat) SetContrast(rhs float32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *PictureFormat) CropBottom() float32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *PictureFormat) SetCropBottom(rhs float32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *PictureFormat) CropLeft() float32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.FltValVal()
}

func (this *PictureFormat) SetCropLeft(rhs float32)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *PictureFormat) CropRight() float32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.FltValVal()
}

func (this *PictureFormat) SetCropRight(rhs float32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *PictureFormat) CropTop() float32 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.FltValVal()
}

func (this *PictureFormat) SetCropTop(rhs float32)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *PictureFormat) TransparencyColor() int32 {
	retVal := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *PictureFormat) SetTransparencyColor(rhs int32)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

func (this *PictureFormat) TransparentBackground() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *PictureFormat) SetTransparentBackground(rhs int32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *PictureFormat) IncrementBrightness(increment float32)  {
	retVal := this.Call(0x0000000a, []interface{}{increment})
	_= retVal
}

func (this *PictureFormat) IncrementContrast(increment float32)  {
	retVal := this.Call(0x0000000b, []interface{}{increment})
	_= retVal
}

func (this *PictureFormat) Crop() *ole.DispatchClass {
	retVal := this.PropGet(0x0000006d, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *PictureFormat) SetCrop(rhs *ole.DispatchClass)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

