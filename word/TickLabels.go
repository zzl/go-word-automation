package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 935D59F5-E365-4F92-B7F5-1C499A63ECA8
var IID_TickLabels = syscall.GUID{0x935D59F5, 0xE365, 0x4F92, 
	[8]byte{0xB7, 0xF5, 0x1C, 0x49, 0x9A, 0x63, 0xEC, 0xA8}}

type TickLabels struct {
	ole.OleClient
}

func NewTickLabels(pDisp *win32.IDispatch, addRef bool, scoped bool) *TickLabels {
	 if pDisp == nil {
		return nil;
	}
	p := &TickLabels{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TickLabelsFromVar(v ole.Variant) *TickLabels {
	return NewTickLabels(v.IDispatch(), false, false)
}

func (this *TickLabels) IID() *syscall.GUID {
	return &IID_TickLabels
}

func (this *TickLabels) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TickLabels) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TickLabels) Delete() ole.Variant {
	retVal, _ := this.Call(0x60020001, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *TickLabels) Font() *ChartFont {
	retVal, _ := this.PropGet(0x60020002, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

func (this *TickLabels) Name() string {
	retVal, _ := this.PropGet(0x60020003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TickLabels) NumberFormat() string {
	retVal, _ := this.PropGet(0x60020004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TickLabels) SetNumberFormat(rhs string)  {
	_ = this.PropPut(0x60020004, []interface{}{rhs})
}

func (this *TickLabels) NumberFormatLinked() bool {
	retVal, _ := this.PropGet(0x60020006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TickLabels) SetNumberFormatLinked(rhs bool)  {
	_ = this.PropPut(0x60020006, []interface{}{rhs})
}

func (this *TickLabels) NumberFormatLocal() ole.Variant {
	retVal, _ := this.PropGet(0x60020008, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *TickLabels) SetNumberFormatLocal(rhs interface{})  {
	_ = this.PropPut(0x60020008, []interface{}{rhs})
}

func (this *TickLabels) Orientation() int32 {
	retVal, _ := this.PropGet(0x6002000a, nil)
	return retVal.LValVal()
}

func (this *TickLabels) SetOrientation(rhs int32)  {
	_ = this.PropPut(0x6002000a, []interface{}{rhs})
}

func (this *TickLabels) Select() ole.Variant {
	retVal, _ := this.Call(0x6002000c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *TickLabels) ReadingOrder() int32 {
	retVal, _ := this.PropGet(0x6002000d, nil)
	return retVal.LValVal()
}

func (this *TickLabels) SetReadingOrder(rhs int32)  {
	_ = this.PropPut(0x6002000d, []interface{}{rhs})
}

func (this *TickLabels) AutoScaleFont() ole.Variant {
	retVal, _ := this.PropGet(0x6002000f, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *TickLabels) SetAutoScaleFont(rhs interface{})  {
	_ = this.PropPut(0x6002000f, []interface{}{rhs})
}

func (this *TickLabels) Depth() int32 {
	retVal, _ := this.PropGet(0x60020011, nil)
	return retVal.LValVal()
}

func (this *TickLabels) Offset() int32 {
	retVal, _ := this.PropGet(0x60020012, nil)
	return retVal.LValVal()
}

func (this *TickLabels) SetOffset(rhs int32)  {
	_ = this.PropPut(0x60020012, []interface{}{rhs})
}

func (this *TickLabels) Alignment() int32 {
	retVal, _ := this.PropGet(0x60020014, nil)
	return retVal.LValVal()
}

func (this *TickLabels) SetAlignment(rhs int32)  {
	_ = this.PropPut(0x60020014, []interface{}{rhs})
}

func (this *TickLabels) MultiLevel() bool {
	retVal, _ := this.PropGet(0x60020016, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TickLabels) SetMultiLevel(rhs bool)  {
	_ = this.PropPut(0x60020016, []interface{}{rhs})
}

func (this *TickLabels) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x60020018, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *TickLabels) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TickLabels) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

