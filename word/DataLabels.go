package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// D8252C5E-EB9F-4D74-AA72-C178B128FAC4
var IID_DataLabels = syscall.GUID{0xD8252C5E, 0xEB9F, 0x4D74, 
	[8]byte{0xAA, 0x72, 0xC1, 0x78, 0xB1, 0x28, 0xFA, 0xC4}}

type DataLabels struct {
	ole.OleClient
}

func NewDataLabels(pDisp *win32.IDispatch, addRef bool, scoped bool) *DataLabels {
	p := &DataLabels{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DataLabelsFromVar(v ole.Variant) *DataLabels {
	return NewDataLabels(v.PdispValVal(), false, false)
}

func (this *DataLabels) IID() *syscall.GUID {
	return &IID_DataLabels
}

func (this *DataLabels) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DataLabels) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DataLabels) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabels) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataLabels) Border() *ChartBorder {
	retVal := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *DataLabels) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataLabels) Interior() *Interior {
	retVal := this.PropGet(0x00000081, nil)
	return NewInterior(retVal.PdispValVal(), false, true)
}

func (this *DataLabels) Fill() *ChartFillFormat {
	retVal := this.PropGet(0x0000067f, nil)
	return NewChartFillFormat(retVal.PdispValVal(), false, true)
}

func (this *DataLabels) Font() *ChartFont {
	retVal := this.PropGet(0x00000092, nil)
	return NewChartFont(retVal.PdispValVal(), false, true)
}

func (this *DataLabels) HorizontalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000088, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataLabels) SetHorizontalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000088, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) Orientation() ole.Variant {
	retVal := this.PropGet(0x00000086, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataLabels) SetOrientation(rhs interface{})  {
	retVal := this.PropPut(0x00000086, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) Shadow() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) VerticalAlignment() ole.Variant {
	retVal := this.PropGet(0x00000089, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataLabels) SetVerticalAlignment(rhs interface{})  {
	retVal := this.PropPut(0x00000089, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) ReadingOrder() int32 {
	retVal := this.PropGet(0x000003cf, nil)
	return retVal.LValVal()
}

func (this *DataLabels) SetReadingOrder(rhs int32)  {
	retVal := this.PropPut(0x000003cf, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) AutoScaleFont() ole.Variant {
	retVal := this.PropGet(0x000005f5, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataLabels) SetAutoScaleFont(rhs interface{})  {
	retVal := this.PropPut(0x000005f5, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) AutoText() bool {
	retVal := this.PropGet(0x00000087, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetAutoText(rhs bool)  {
	retVal := this.PropPut(0x00000087, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) NumberFormat() string {
	retVal := this.PropGet(0x000000c1, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DataLabels) SetNumberFormat(rhs string)  {
	retVal := this.PropPut(0x000000c1, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) NumberFormatLinked() bool {
	retVal := this.PropGet(0x000000c2, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetNumberFormatLinked(rhs bool)  {
	retVal := this.PropPut(0x000000c2, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) NumberFormatLocal() ole.Variant {
	retVal := this.PropGet(0x00000449, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataLabels) SetNumberFormatLocal(rhs interface{})  {
	retVal := this.PropPut(0x00000449, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) ShowLegendKey() bool {
	retVal := this.PropGet(0x000000ab, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowLegendKey(rhs bool)  {
	retVal := this.PropPut(0x000000ab, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) Type() ole.Variant {
	retVal := this.PropGet(0x0000006c, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataLabels) SetType(rhs interface{})  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) Position() int32 {
	retVal := this.PropGet(0x00000085, nil)
	return retVal.LValVal()
}

func (this *DataLabels) SetPosition(rhs int32)  {
	retVal := this.PropPut(0x00000085, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) ShowSeriesName() bool {
	retVal := this.PropGet(0x000007e6, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowSeriesName(rhs bool)  {
	retVal := this.PropPut(0x000007e6, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) ShowCategoryName() bool {
	retVal := this.PropGet(0x000007e7, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowCategoryName(rhs bool)  {
	retVal := this.PropPut(0x000007e7, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) ShowValue() bool {
	retVal := this.PropGet(0x000007e8, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowValue(rhs bool)  {
	retVal := this.PropPut(0x000007e8, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) ShowPercentage() bool {
	retVal := this.PropGet(0x000007e9, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowPercentage(rhs bool)  {
	retVal := this.PropPut(0x000007e9, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) ShowBubbleSize() bool {
	retVal := this.PropGet(0x000007ea, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DataLabels) SetShowBubbleSize(rhs bool)  {
	retVal := this.PropPut(0x000007ea, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) Separator() ole.Variant {
	retVal := this.PropGet(0x000007eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *DataLabels) SetSeparator(rhs interface{})  {
	retVal := this.PropPut(0x000007eb, []interface{}{rhs})
	_= retVal
}

func (this *DataLabels) Count() int32 {
	retVal := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *DataLabels) Item(index interface{}) *DataLabel {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewDataLabel(retVal.PdispValVal(), false, true)
}

func (this *DataLabels) NewEnum_() *com.UnknownClass {
	retVal := this.Call(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *DataLabels) ForEach(action func(item *DataLabel) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*DataLabel)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *DataLabels) Format() *ChartFormat {
	retVal := this.PropGet(0x60020032, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *DataLabels) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DataLabels) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *DataLabels) Default_(index interface{}) *DataLabel {
	retVal := this.Call(0x60020035, []interface{}{index})
	return NewDataLabel(retVal.PdispValVal(), false, true)
}

