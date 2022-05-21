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
	return NewField(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Field) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Field) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Field) Code() *Range {
	retVal := this.PropGet(0x00000000, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Field) SetCode(rhs *Range)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *Field) Type() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Field) Locked() bool {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Field) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *Field) Kind() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Field) Result() *Range {
	retVal := this.PropGet(0x00000004, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Field) SetResult(rhs *Range)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *Field) Data() string {
	retVal := this.PropGet(0x00000005, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Field) SetData(rhs string)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *Field) Next() *Field {
	retVal := this.PropGet(0x00000006, nil)
	return NewField(retVal.PdispValVal(), false, true)
}

func (this *Field) Previous() *Field {
	retVal := this.PropGet(0x00000007, nil)
	return NewField(retVal.PdispValVal(), false, true)
}

func (this *Field) Index() int32 {
	retVal := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Field) ShowCodes() bool {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Field) SetShowCodes(rhs bool)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *Field) LinkFormat() *LinkFormat {
	retVal := this.PropGet(0x0000000a, nil)
	return NewLinkFormat(retVal.PdispValVal(), false, true)
}

func (this *Field) OLEFormat() *OLEFormat {
	retVal := this.PropGet(0x0000000b, nil)
	return NewOLEFormat(retVal.PdispValVal(), false, true)
}

func (this *Field) InlineShape() *InlineShape {
	retVal := this.PropGet(0x0000000c, nil)
	return NewInlineShape(retVal.PdispValVal(), false, true)
}

func (this *Field) Select()  {
	retVal := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *Field) Update() bool {
	retVal := this.Call(0x00000065, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Field) Unlink()  {
	retVal := this.Call(0x00000066, nil)
	_= retVal
}

func (this *Field) UpdateSource()  {
	retVal := this.Call(0x00000067, nil)
	_= retVal
}

func (this *Field) DoClick()  {
	retVal := this.Call(0x00000068, nil)
	_= retVal
}

func (this *Field) Copy()  {
	retVal := this.Call(0x00000069, nil)
	_= retVal
}

func (this *Field) Cut()  {
	retVal := this.Call(0x0000006a, nil)
	_= retVal
}

func (this *Field) Delete()  {
	retVal := this.Call(0x0000006b, nil)
	_= retVal
}

