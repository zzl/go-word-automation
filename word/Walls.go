package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// A2E94180-7564-4D97-806B-BBC0D0A1350C
var IID_Walls = syscall.GUID{0xA2E94180, 0x7564, 0x4D97, 
	[8]byte{0x80, 0x6B, 0xBB, 0xC0, 0xD0, 0xA1, 0x35, 0x0C}}

type Walls struct {
	ole.OleClient
}

func NewWalls(pDisp *win32.IDispatch, addRef bool, scoped bool) *Walls {
	 if pDisp == nil {
		return nil;
	}
	p := &Walls{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WallsFromVar(v ole.Variant) *Walls {
	return NewWalls(v.IDispatch(), false, false)
}

func (this *Walls) IID() *syscall.GUID {
	return &IID_Walls
}

func (this *Walls) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Walls) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Walls) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Walls) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Walls) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *Walls) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Walls) Interior() *Interior {
	retVal, _ := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.IDispatch(), false, true)
}

func (this *Walls) Fill() *ChartFillFormat {
	retVal, _ := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.IDispatch(), false, true)
}

func (this *Walls) PictureType() ole.Variant {
	retVal, _ := this.PropGet(0x000000a1, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Walls) SetPictureType(rhs interface{})  {
	_ = this.PropPut(0x000000a1, []interface{}{rhs})
}

func (this *Walls) Paste()  {
	retVal, _ := this.Call(0x000000d3, nil)
	_= retVal
}

func (this *Walls) PictureUnit() ole.Variant {
	retVal, _ := this.PropGet(0x000000a2, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Walls) SetPictureUnit(rhs interface{})  {
	_ = this.PropPut(0x000000a2, []interface{}{rhs})
}

func (this *Walls) Thickness() int32 {
	retVal, _ := this.PropGet(0x00000973, nil)
	return retVal.LValVal()
}

func (this *Walls) SetThickness(rhs int32)  {
	_ = this.PropPut(0x00000973, []interface{}{rhs})
}

func (this *Walls) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x6002000e, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *Walls) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Walls) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

