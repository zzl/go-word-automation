package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// DD8F80B8-9B80-4E89-9BEC-F12DF35E43B3
var IID_ChartColorFormat = syscall.GUID{0xDD8F80B8, 0x9B80, 0x4E89, 
	[8]byte{0x9B, 0xEC, 0xF1, 0x2D, 0xF3, 0x5E, 0x43, 0xB3}}

type ChartColorFormat struct {
	ole.OleClient
}

func NewChartColorFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartColorFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &ChartColorFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartColorFormatFromVar(v ole.Variant) *ChartColorFormat {
	return NewChartColorFormat(v.IDispatch(), false, false)
}

func (this *ChartColorFormat) IID() *syscall.GUID {
	return &IID_ChartColorFormat
}

func (this *ChartColorFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartColorFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartColorFormat) SchemeColor() int32 {
	retVal, _ := this.PropGet(0x0000066e, nil)
	return retVal.LValVal()
}

func (this *ChartColorFormat) SetSchemeColor(rhs int32)  {
	_ = this.PropPut(0x0000066e, []interface{}{rhs})
}

func (this *ChartColorFormat) RGB() int32 {
	retVal, _ := this.PropGet(0x0000041f, nil)
	return retVal.LValVal()
}

func (this *ChartColorFormat) Default_() int32 {
	retVal, _ := this.PropGet(0x60020005, nil)
	return retVal.LValVal()
}

func (this *ChartColorFormat) Type() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *ChartColorFormat) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartColorFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

