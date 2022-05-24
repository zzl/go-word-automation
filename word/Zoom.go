package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209A6-0000-0000-C000-000000000046
var IID_Zoom = syscall.GUID{0x000209A6, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Zoom struct {
	ole.OleClient
}

func NewZoom(pDisp *win32.IDispatch, addRef bool, scoped bool) *Zoom {
	 if pDisp == nil {
		return nil;
	}
	p := &Zoom{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ZoomFromVar(v ole.Variant) *Zoom {
	return NewZoom(v.IDispatch(), false, false)
}

func (this *Zoom) IID() *syscall.GUID {
	return &IID_Zoom
}

func (this *Zoom) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Zoom) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Zoom) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Zoom) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Zoom) Percentage() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *Zoom) SetPercentage(rhs int32)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *Zoom) PageFit() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Zoom) SetPageFit(rhs int32)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *Zoom) PageRows() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Zoom) SetPageRows(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *Zoom) PageColumns() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Zoom) SetPageColumns(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

