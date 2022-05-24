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
	 if pDisp == nil {
		return nil;
	}
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
	return NewDataTable(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x60020000, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetShowLegendKey(rhs bool)  {
	_ = this.PropPut(0x60020000, []interface{}{rhs})
}

func (this *DataTable) HasBorderHorizontal() bool {
	retVal, _ := this.PropGet(0x60020002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetHasBorderHorizontal(rhs bool)  {
	_ = this.PropPut(0x60020002, []interface{}{rhs})
}

func (this *DataTable) HasBorderVertical() bool {
	retVal, _ := this.PropGet(0x60020004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetHasBorderVertical(rhs bool)  {
	_ = this.PropPut(0x60020004, []interface{}{rhs})
}

func (this *DataTable) HasBorderOutline() bool {
	retVal, _ := this.PropGet(0x60020006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataTable) SetHasBorderOutline(rhs bool)  {
	_ = this.PropPut(0x60020006, []interface{}{rhs})
}

func (this *DataTable) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x60020008, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *DataTable) Font() *ChartFont {
	retVal, _ := this.PropGet(0x60020009, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

func (this *DataTable) Select()  {
	retVal, _ := this.Call(0x6002000a, nil)
	_= retVal
}

func (this *DataTable) Delete()  {
	retVal, _ := this.Call(0x6002000b, nil)
	_= retVal
}

func (this *DataTable) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x6002000c, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DataTable) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x6002000d, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *DataTable) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x6002000d, []interface{}{rhs})
}

func (this *DataTable) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x6002000f, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *DataTable) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DataTable) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

