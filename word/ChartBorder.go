package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// AB0D33A3-C9EA-485B-9443-4C1BB3656CEA
var IID_ChartBorder = syscall.GUID{0xAB0D33A3, 0xC9EA, 0x485B, 
	[8]byte{0x94, 0x43, 0x4C, 0x1B, 0xB3, 0x65, 0x6C, 0xEA}}

type ChartBorder struct {
	ole.OleClient
}

func NewChartBorder(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartBorder {
	p := &ChartBorder{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartBorderFromVar(v ole.Variant) *ChartBorder {
	return NewChartBorder(v.PdispValVal(), false, false)
}

func (this *ChartBorder) IID() *syscall.GUID {
	return &IID_ChartBorder
}

func (this *ChartBorder) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartBorder) Color() ole.Variant {
	retVal := this.PropGet(0x60020000, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartBorder) SetColor(rhs interface{})  {
	retVal := this.PropPut(0x60020000, []interface{}{rhs})
	_= retVal
}

func (this *ChartBorder) ColorIndex() ole.Variant {
	retVal := this.PropGet(0x60020002, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartBorder) SetColorIndex(rhs interface{})  {
	retVal := this.PropPut(0x60020002, []interface{}{rhs})
	_= retVal
}

func (this *ChartBorder) LineStyle() ole.Variant {
	retVal := this.PropGet(0x60020004, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartBorder) SetLineStyle(rhs interface{})  {
	retVal := this.PropPut(0x60020004, []interface{}{rhs})
	_= retVal
}

func (this *ChartBorder) Weight() ole.Variant {
	retVal := this.PropGet(0x60020006, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartBorder) SetWeight(rhs interface{})  {
	retVal := this.PropPut(0x60020006, []interface{}{rhs})
	_= retVal
}

func (this *ChartBorder) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartBorder) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartBorder) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

