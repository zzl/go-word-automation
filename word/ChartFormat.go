package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// B66D3C1A-4541-4961-B35B-A353C03F6A99
var IID_ChartFormat = syscall.GUID{0xB66D3C1A, 0x4541, 0x4961, 
	[8]byte{0xB3, 0x5B, 0xA3, 0x53, 0xC0, 0x3F, 0x6A, 0x99}}

type ChartFormat struct {
	ole.OleClient
}

func NewChartFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartFormat {
	p := &ChartFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartFormatFromVar(v ole.Variant) *ChartFormat {
	return NewChartFormat(v.PdispValVal(), false, false)
}

func (this *ChartFormat) IID() *syscall.GUID {
	return &IID_ChartFormat
}

func (this *ChartFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartFormat) Fill() *FillFormat {
	retVal := this.PropGet(0x60020000, nil)
	return NewFillFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) Glow() *GlowFormat {
	retVal := this.PropGet(0x60020001, nil)
	return NewGlowFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) Line() *LineFormat {
	retVal := this.PropGet(0x60020002, nil)
	return NewLineFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x60020003, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartFormat) PictureFormat() *PictureFormat {
	retVal := this.PropGet(0x60020004, nil)
	return NewPictureFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) Shadow() *ShadowFormat {
	retVal := this.PropGet(0x60020005, nil)
	return NewShadowFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) SoftEdge() *SoftEdgeFormat {
	retVal := this.PropGet(0x60020006, nil)
	return NewSoftEdgeFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) TextFrame2() *ole.DispatchClass {
	retVal := this.PropGet(0x60020007, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartFormat) ThreeD() *ThreeDFormat {
	retVal := this.PropGet(0x60020008, nil)
	return NewThreeDFormat(retVal.PdispValVal(), false, true)
}

func (this *ChartFormat) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartFormat) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

