package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 02B17CB4-7D55-4B34-B38B-10381433441F
var IID_OMathGroupChar = syscall.GUID{0x02B17CB4, 0x7D55, 0x4B34, 
	[8]byte{0xB3, 0x8B, 0x10, 0x38, 0x14, 0x33, 0x44, 0x1F}}

type OMathGroupChar struct {
	ole.OleClient
}

func NewOMathGroupChar(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathGroupChar {
	p := &OMathGroupChar{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathGroupCharFromVar(v ole.Variant) *OMathGroupChar {
	return NewOMathGroupChar(v.PdispValVal(), false, false)
}

func (this *OMathGroupChar) IID() *syscall.GUID {
	return &IID_OMathGroupChar
}

func (this *OMathGroupChar) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathGroupChar) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathGroupChar) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathGroupChar) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathGroupChar) E() *OMath {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathGroupChar) Char() int16 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.IValVal()
}

func (this *OMathGroupChar) SetChar(rhs int16)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *OMathGroupChar) CharTop() bool {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathGroupChar) SetCharTop(rhs bool)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *OMathGroupChar) AlignTop() bool {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathGroupChar) SetAlignTop(rhs bool)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

