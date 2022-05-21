package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// C94688A6-A2A7-4133-A26D-726CD569D5F3
var IID_OMathDelim = syscall.GUID{0xC94688A6, 0xA2A7, 0x4133, 
	[8]byte{0xA2, 0x6D, 0x72, 0x6C, 0xD5, 0x69, 0xD5, 0xF3}}

type OMathDelim struct {
	ole.OleClient
}

func NewOMathDelim(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathDelim {
	p := &OMathDelim{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathDelimFromVar(v ole.Variant) *OMathDelim {
	return NewOMathDelim(v.PdispValVal(), false, false)
}

func (this *OMathDelim) IID() *syscall.GUID {
	return &IID_OMathDelim
}

func (this *OMathDelim) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathDelim) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathDelim) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathDelim) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathDelim) E() *OMathArgs {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMathArgs(retVal.PdispValVal(), false, true)
}

func (this *OMathDelim) BegChar() int16 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.IValVal()
}

func (this *OMathDelim) SetBegChar(rhs int16)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *OMathDelim) SepChar() int16 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.IValVal()
}

func (this *OMathDelim) SetSepChar(rhs int16)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *OMathDelim) EndChar() int16 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.IValVal()
}

func (this *OMathDelim) SetEndChar(rhs int16)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *OMathDelim) Grow() bool {
	retVal := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathDelim) SetGrow(rhs bool)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

func (this *OMathDelim) Shape() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *OMathDelim) SetShape(rhs int32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *OMathDelim) NoLeftChar() bool {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathDelim) SetNoLeftChar(rhs bool)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *OMathDelim) NoRightChar() bool {
	retVal := this.PropGet(0x0000006e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathDelim) SetNoRightChar(rhs bool)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

