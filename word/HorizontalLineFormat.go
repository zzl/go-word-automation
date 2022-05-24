package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209DE-0000-0000-C000-000000000046
var IID_HorizontalLineFormat = syscall.GUID{0x000209DE, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HorizontalLineFormat struct {
	ole.OleClient
}

func NewHorizontalLineFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *HorizontalLineFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &HorizontalLineFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HorizontalLineFormatFromVar(v ole.Variant) *HorizontalLineFormat {
	return NewHorizontalLineFormat(v.IDispatch(), false, false)
}

func (this *HorizontalLineFormat) IID() *syscall.GUID {
	return &IID_HorizontalLineFormat
}

func (this *HorizontalLineFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HorizontalLineFormat) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *HorizontalLineFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HorizontalLineFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *HorizontalLineFormat) PercentWidth() float32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.FltValVal()
}

func (this *HorizontalLineFormat) SetPercentWidth(rhs float32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *HorizontalLineFormat) NoShade() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *HorizontalLineFormat) SetNoShade(rhs bool)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *HorizontalLineFormat) Alignment() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *HorizontalLineFormat) SetAlignment(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *HorizontalLineFormat) WidthType() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *HorizontalLineFormat) SetWidthType(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

