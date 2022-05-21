package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020915-0000-0000-C000-000000000046
var IID_CustomLabel = syscall.GUID{0x00020915, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CustomLabel struct {
	ole.OleClient
}

func NewCustomLabel(pDisp *win32.IDispatch, addRef bool, scoped bool) *CustomLabel {
	p := &CustomLabel{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CustomLabelFromVar(v ole.Variant) *CustomLabel {
	return NewCustomLabel(v.PdispValVal(), false, false)
}

func (this *CustomLabel) IID() *syscall.GUID {
	return &IID_CustomLabel
}

func (this *CustomLabel) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CustomLabel) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CustomLabel) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CustomLabel) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CustomLabel) Index() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *CustomLabel) Name() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CustomLabel) SetName(rhs string)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) TopMargin() float32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.FltValVal()
}

func (this *CustomLabel) SetTopMargin(rhs float32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) SideMargin() float32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.FltValVal()
}

func (this *CustomLabel) SetSideMargin(rhs float32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) Height() float32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.FltValVal()
}

func (this *CustomLabel) SetHeight(rhs float32)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) Width() float32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.FltValVal()
}

func (this *CustomLabel) SetWidth(rhs float32)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) VerticalPitch() float32 {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.FltValVal()
}

func (this *CustomLabel) SetVerticalPitch(rhs float32)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) HorizontalPitch() float32 {
	retVal := this.PropGet(0x00000008, nil)
	return retVal.FltValVal()
}

func (this *CustomLabel) SetHorizontalPitch(rhs float32)  {
	retVal := this.PropPut(0x00000008, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) NumberAcross() int32 {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.LValVal()
}

func (this *CustomLabel) SetNumberAcross(rhs int32)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) NumberDown() int32 {
	retVal := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *CustomLabel) SetNumberDown(rhs int32)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) DotMatrix() bool {
	retVal := this.PropGet(0x0000000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CustomLabel) PageSize() int32 {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *CustomLabel) SetPageSize(rhs int32)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *CustomLabel) Valid() bool {
	retVal := this.PropGet(0x0000000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CustomLabel) Delete()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

