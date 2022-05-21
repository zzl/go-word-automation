package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// F01943FF-1985-445E-8602-8FB8F39CCA75
var IID_ReflectionFormat = syscall.GUID{0xF01943FF, 0x1985, 0x445E, 
	[8]byte{0x86, 0x02, 0x8F, 0xB8, 0xF3, 0x9C, 0xCA, 0x75}}

type ReflectionFormat struct {
	ole.OleClient
}

func NewReflectionFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ReflectionFormat {
	p := &ReflectionFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ReflectionFormatFromVar(v ole.Variant) *ReflectionFormat {
	return NewReflectionFormat(v.PdispValVal(), false, false)
}

func (this *ReflectionFormat) IID() *syscall.GUID {
	return &IID_ReflectionFormat
}

func (this *ReflectionFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ReflectionFormat) Type() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *ReflectionFormat) SetType(rhs int32)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *ReflectionFormat) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ReflectionFormat) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ReflectionFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ReflectionFormat) Transparency() float32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.FltValVal()
}

func (this *ReflectionFormat) SetTransparency(rhs float32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *ReflectionFormat) Size() float32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.FltValVal()
}

func (this *ReflectionFormat) SetSize(rhs float32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *ReflectionFormat) Offset() float32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.FltValVal()
}

func (this *ReflectionFormat) SetOffset(rhs float32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *ReflectionFormat) Blur() float32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.FltValVal()
}

func (this *ReflectionFormat) SetBlur(rhs float32)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

