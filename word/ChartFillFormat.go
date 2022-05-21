package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// F152D349-7D20-4C01-A42B-2D6DE4F3891C
var IID_ChartFillFormat = syscall.GUID{0xF152D349, 0x7D20, 0x4C01, 
	[8]byte{0xA4, 0x2B, 0x2D, 0x6D, 0xE4, 0xF3, 0x89, 0x1C}}

type ChartFillFormat struct {
	ole.OleClient
}

func NewChartFillFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartFillFormat {
	p := &ChartFillFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartFillFormatFromVar(v ole.Variant) *ChartFillFormat {
	return NewChartFillFormat(v.PdispValVal(), false, false)
}

func (this *ChartFillFormat) IID() *syscall.GUID {
	return &IID_ChartFillFormat
}

func (this *ChartFillFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartFillFormat) OneColorGradient(style int32, variant int32, degree float32)  {
	retVal := this.Call(0x60020000, []interface{}{style, variant, degree})
	_= retVal
}

func (this *ChartFillFormat) TwoColorGradient(style int32, variant int32)  {
	retVal := this.Call(0x60020001, []interface{}{style, variant})
	_= retVal
}

func (this *ChartFillFormat) PresetTextured(presetTexture int32)  {
	retVal := this.Call(0x60020002, []interface{}{presetTexture})
	_= retVal
}

func (this *ChartFillFormat) Solid()  {
	retVal := this.Call(0x60020003, nil)
	_= retVal
}

func (this *ChartFillFormat) Patterned(pattern int32)  {
	retVal := this.Call(0x60020004, []interface{}{pattern})
	_= retVal
}

var ChartFillFormat_UserPicture_OptArgs= []string{
	"PictureFile", "PictureFormat", "PictureStackUnit", "PicturePlacement", 
}

func (this *ChartFillFormat) UserPicture(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(ChartFillFormat_UserPicture_OptArgs, optArgs)
	retVal := this.Call(0x60020005, nil, optArgs...)
	_= retVal
}

func (this *ChartFillFormat) UserTextured(textureFile string)  {
	retVal := this.Call(0x60020006, []interface{}{textureFile})
	_= retVal
}

func (this *ChartFillFormat) PresetGradient(style int32, variant int32, presetGradientType int32)  {
	retVal := this.Call(0x60020007, []interface{}{style, variant, presetGradientType})
	_= retVal
}

func (this *ChartFillFormat) BackColor() *ChartColorFormat {
	retVal := this.PropGet(0x60020008, nil)
	return NewChartColorFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFillFormat) ForeColor() *ChartColorFormat {
	retVal := this.PropGet(0x60020009, nil)
	return NewChartColorFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFillFormat) GradientColorType() int32 {
	retVal := this.PropGet(0x6002000a, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) GradientDegree() float32 {
	retVal := this.PropGet(0x6002000b, nil)
	return retVal.FltValVal()
}

func (this *ChartFillFormat) GradientStyle() int32 {
	retVal := this.PropGet(0x6002000c, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) GradientVariant() int32 {
	retVal := this.PropGet(0x6002000d, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) Pattern() int32 {
	retVal := this.PropGet(0x6002000e, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) PresetGradientType() int32 {
	retVal := this.PropGet(0x6002000f, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) PresetTexture() int32 {
	retVal := this.PropGet(0x60020010, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) TextureName() string {
	retVal := this.PropGet(0x60020011, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartFillFormat) TextureType() int32 {
	retVal := this.PropGet(0x60020012, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) Type() int32 {
	retVal := this.PropGet(0x60020013, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) Visible() int32 {
	retVal := this.PropGet(0x60020014, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) SetVisible(rhs int32)  {
	retVal := this.PropPut(0x60020014, []interface{}{rhs})
	_= retVal
}

func (this *ChartFillFormat) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartFillFormat) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartFillFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

