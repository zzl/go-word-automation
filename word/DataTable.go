package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// DCE9F2C4-4C02-43BA-840E-B4276550EF79
var IID_DataTable = syscall.GUID{0xDCE9F2C4, 0x4C02, 0x43BA, 
	[8]byte{0x84, 0x0E, 0xB4, 0x27, 0x65, 0x50, 0xEF, 0x79}}

type DataTable struct {
	ole.OleClient
}

func NewDataTable(pDisp *win32.IDispatch, addRef bool, scoped bool) *DataTable {
	p := &DataTable{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DataTableFromVar(v ole.Variant) *DataTable {
	return NewDataTable(v.PdispValVal(), false, false)
}

func (this *DataTable) IID() *syscall.GUID {
	return &IID_DataTable
}

func (this *DataTable) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DataTable) ShowLegendKey() bool {
	retVal := this.PropGet(0x60020000, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetShowLegendKey(rhs bool)  {
	retVal := this.PropPut(0x60020000, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) HasBorderHorizontal() bool {
	retVal := this.PropGet(0x60020002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetHasBorderHorizontal(rhs bool)  {
	retVal := this.PropPut(0x60020002, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) HasBorderVertical() bool {
	retVal := this.PropGet(0x60020004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetHasBorderVertical(rhs bool)  {
	retVal := this.PropPut(0x60020004, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) HasBorderOutline() bool {
	retVal := this.PropGet(0x60020006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetHasBorderOutline(rhs bool)  {
	retVal := this.PropPut(0x60020006, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) Border() *ChartBorder {
	retVal := this.PropGet(0x60020008, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *DataTable) Font() *ChartFont {
	retVal := this.PropGet(0x60020009, nil)
	return NewChartFont(retVal.PdispValVal(), false, true)
}

func (this *DataTable) Select()  {
	retVal := this.Call(0x6002000a, nil)
	_= retVal
}

func (this *DataTable) Delete()  {
	retVal := this.Call(0x6002000b, nil)
	_= retVal
}

func (this *DataTable) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x6002000c, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DataTable) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x6002000d, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataTable) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x6002000d, []interface{}{rhs})
	_= retVal
}

func (this *DataTable) Format() *ChartFormat {
	retVal := this.PropGet(0x6002000f, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *DataTable) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DataTable) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

