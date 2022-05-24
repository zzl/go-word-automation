package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209C6-0000-0000-C000-000000000046
var IID_ColorFormat = syscall.GUID{0x000209C6, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ColorFormat struct {
	ole.OleClient
}

func NewColorFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ColorFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &ColorFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ColorFormatFromVar(v ole.Variant) *ColorFormat {
	return NewColorFormat(v.IDispatch(), false, false)
}

func (this *ColorFormat) IID() *syscall.GUID {
	return &IID_ColorFormat
}

func (this *ColorFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ColorFormat) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ColorFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ColorFormat) RGB() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetRGB(rhs int32)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *ColorFormat) SchemeColor() int32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetSchemeColor(rhs int32)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *ColorFormat) Type() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) Name() string {
	retVal, _ := this.PropGet(0x00000066, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ColorFormat) SetName(rhs string)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *ColorFormat) TintAndShade() float32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *ColorFormat) SetTintAndShade(rhs float32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *ColorFormat) OverPrint() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetOverPrint(rhs int32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *ColorFormat) Ink(index int32) float32 {
	retVal, _ := this.PropGet(0x00000069, []interface{}{index})
	return retVal.FltValVal()
}

func (this *ColorFormat) SetInk(index int32, rhs float32)  {
	_ = this.PropPut(0x00000069, []interface{}{index, rhs})
}

func (this *ColorFormat) Cyan() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetCyan(rhs int32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *ColorFormat) Magenta() int32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetMagenta(rhs int32)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *ColorFormat) Yellow() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetYellow(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *ColorFormat) Black() int32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetBlack(rhs int32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *ColorFormat) SetCMYK(cyan int32, magenta int32, yellow int32, black int32)  {
	retVal, _ := this.Call(0x0000006e, []interface{}{cyan, magenta, yellow, black})
	_= retVal
}

func (this *ColorFormat) ObjectThemeColor() int32 {
	retVal, _ := this.PropGet(0x000000c8, nil)
	return retVal.LValVal()
}

func (this *ColorFormat) SetObjectThemeColor(rhs int32)  {
	_ = this.PropPut(0x000000c8, []interface{}{rhs})
}

func (this *ColorFormat) Brightness() float32 {
	retVal, _ := this.PropGet(0x000000c9, nil)
	return retVal.FltValVal()
}

func (this *ColorFormat) SetBrightness(rhs float32)  {
	_ = this.PropPut(0x000000c9, []interface{}{rhs})
}

