package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002093A-0000-0000-C000-000000000046
var IID_Shading = syscall.GUID{0x0002093A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Shading struct {
	ole.OleClient
}

func NewShading(pDisp *win32.IDispatch, addRef bool, scoped bool) *Shading {
	 if pDisp == nil {
		return nil;
	}
	p := &Shading{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ShadingFromVar(v ole.Variant) *Shading {
	return NewShading(v.IDispatch(), false, false)
}

func (this *Shading) IID() *syscall.GUID {
	return &IID_Shading
}

func (this *Shading) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Shading) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Shading) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Shading) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Shading) ForegroundPatternColorIndex() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Shading) SetForegroundPatternColorIndex(rhs int32)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *Shading) BackgroundPatternColorIndex() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Shading) SetBackgroundPatternColorIndex(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *Shading) Texture() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Shading) SetTexture(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Shading) ForegroundPatternColor() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Shading) SetForegroundPatternColor(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *Shading) BackgroundPatternColor() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Shading) SetBackgroundPatternColor(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

