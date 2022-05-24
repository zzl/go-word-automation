package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002092F-0000-0000-C000-000000000046
var IID_Field = syscall.GUID{0x0002092F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Field struct {
	ole.OleClient
}

func NewField(pDisp *win32.IDispatch, addRef bool, scoped bool) *Field {
	 if pDisp == nil {
		return nil;
	}
	p := &Field{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FieldFromVar(v ole.Variant) *Field {
	return NewField(v.IDispatch(), false, false)
}

func (this *Field) IID() *syscall.GUID {
	return &IID_Field
}

func (this *Field) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Field) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Field) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Field) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Field) Code() *Range {
	retVal, _ := this.PropGet(0x00000000, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Field) SetCode(rhs *Range)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *Field) Type() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Field) Locked() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Field) SetLocked(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *Field) Kind() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Field) Result() *Range {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Field) SetResult(rhs *Range)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *Field) Data() string {
	retVal, _ := this.PropGet(0x00000005, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Field) SetData(rhs string)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *Field) Next() *Field {
	retVal, _ := this.PropGet(0x00000006, nil)
	return NewField(retVal.IDispatch(), false, true)
}

func (this *Field) Previous() *Field {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewField(retVal.IDispatch(), false, true)
}

func (this *Field) Index() int32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Field) ShowCodes() bool {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Field) SetShowCodes(rhs bool)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *Field) LinkFormat() *LinkFormat {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewLinkFormat(retVal.IDispatch(), false, true)
}

func (this *Field) OLEFormat() *OLEFormat {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return NewOLEFormat(retVal.IDispatch(), false, true)
}

func (this *Field) InlineShape() *InlineShape {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

func (this *Field) Select()  {
	retVal, _ := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *Field) Update() bool {
	retVal, _ := this.Call(0x00000065, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Field) Unlink()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *Field) UpdateSource()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

func (this *Field) DoClick()  {
	retVal, _ := this.Call(0x00000068, nil)
	_= retVal
}

func (this *Field) Copy()  {
	retVal, _ := this.Call(0x00000069, nil)
	_= retVal
}

func (this *Field) Cut()  {
	retVal, _ := this.Call(0x0000006a, nil)
	_= retVal
}

func (this *Field) Delete()  {
	retVal, _ := this.Call(0x0000006b, nil)
	_= retVal
}

