package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// F1B14F40-5C32-4C8C-B5B2-DE537BB6B89D
var IID_GlowFormat = syscall.GUID{0xF1B14F40, 0x5C32, 0x4C8C, 
	[8]byte{0xB5, 0xB2, 0xDE, 0x53, 0x7B, 0xB6, 0xB8, 0x9D}}

type GlowFormat struct {
	ole.OleClient
}

func NewGlowFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *GlowFormat {
	 if pDisp == nil {
		return nil;
	}
	p := &GlowFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func GlowFormatFromVar(v ole.Variant) *GlowFormat {
	return NewGlowFormat(v.IDispatch(), false, false)
}

func (this *GlowFormat) IID() *syscall.GUID {
	return &IID_GlowFormat
}

func (this *GlowFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *GlowFormat) Radius() float32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.FltValVal()
}

func (this *GlowFormat) SetRadius(rhs float32)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *GlowFormat) Color() *ColorFormat {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewColorFormat(retVal.IDispatch(), false, true)
}

func (this *GlowFormat) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *GlowFormat) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *GlowFormat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *GlowFormat) Transparency() float32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.FltValVal()
}

func (this *GlowFormat) SetTransparency(rhs float32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

