package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 7E64D2BE-2818-48CB-8F8A-CC7B61D9E860
var IID_Floor = syscall.GUID{0x7E64D2BE, 0x2818, 0x48CB, 
	[8]byte{0x8F, 0x8A, 0xCC, 0x7B, 0x61, 0xD9, 0xE8, 0x60}}

type Floor struct {
	ole.OleClient
}

func NewFloor(pDisp *win32.IDispatch, addRef bool, scoped bool) *Floor {
	p := &Floor{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FloorFromVar(v ole.Variant) *Floor {
	return NewFloor(v.PdispValVal(), false, false)
}

func (this *Floor) IID() *syscall.GUID {
	return &IID_Floor
}

func (this *Floor) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Floor) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Floor) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Floor) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Floor) Border() *ChartBorder {
	retVal := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *Floor) ClearFormats() ole.Variant {
	retVal := this.Call(0x00000070, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Floor) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *Floor) Fill() *ChartFillFormat {
	retVal := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.PdispValVal(), false, true)
}

func (this *Floor) PictureType() ole.Variant {
	retVal := this.PropGet(0x000000a1, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Floor) SetPictureType(rhs interface{})  {
	retVal := this.PropPut(0x000000a1, []interface{}{rhs})
	_= retVal
}

func (this *Floor) Paste()  {
	retVal := this.Call(0x000000d3, nil)
	_= retVal
}

func (this *Floor) Thickness() int32 {
	retVal := this.PropGet(0x00000973, nil)
	return retVal.LValVal()
}

func (this *Floor) SetThickness(rhs int32)  {
	retVal := this.PropPut(0x00000973, []interface{}{rhs})
	_= retVal
}

func (this *Floor) Format() *ChartFormat {
	retVal := this.PropGet(0x6002000c, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *Floor) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Floor) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

