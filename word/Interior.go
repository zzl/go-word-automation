package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// B184502B-587A-4C6A-8DC4-ECE4354883C6
var IID_Interior = syscall.GUID{0xB184502B, 0x587A, 0x4C6A, 
	[8]byte{0x8D, 0xC4, 0xEC, 0xE4, 0x35, 0x48, 0x83, 0xC6}}

type Interior struct {
	ole.OleClient
}

func NewInterior(pDisp *win32.IDispatch, addRef bool, scoped bool) *Interior {
	p := &Interior{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func InteriorFromVar(v ole.Variant) *Interior {
	return NewInterior(v.PdispValVal(), false, false)
}

func (this *Interior) IID() *syscall.GUID {
	return &IID_Interior
}

func (this *Interior) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Interior) Color() ole.Variant {
	retVal := this.PropGet(0x60020000, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetColor(rhs interface{})  {
	retVal := this.PropPut(0x60020000, []interface{}{rhs})
	_= retVal
}

func (this *Interior) ColorIndex() ole.Variant {
	retVal := this.PropGet(0x60020002, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetColorIndex(rhs interface{})  {
	retVal := this.PropPut(0x60020002, []interface{}{rhs})
	_= retVal
}

func (this *Interior) InvertIfNegative() ole.Variant {
	retVal := this.PropGet(0x60020004, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetInvertIfNegative(rhs interface{})  {
	retVal := this.PropPut(0x60020004, []interface{}{rhs})
	_= retVal
}

func (this *Interior) Pattern() ole.Variant {
	retVal := this.PropGet(0x60020006, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetPattern(rhs interface{})  {
	retVal := this.PropPut(0x60020006, []interface{}{rhs})
	_= retVal
}

func (this *Interior) PatternColor() ole.Variant {
	retVal := this.PropGet(0x60020008, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetPatternColor(rhs interface{})  {
	retVal := this.PropPut(0x60020008, []interface{}{rhs})
	_= retVal
}

func (this *Interior) PatternColorIndex() ole.Variant {
	retVal := this.PropGet(0x6002000a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Interior) SetPatternColorIndex(rhs interface{})  {
	retVal := this.PropPut(0x6002000a, []interface{}{rhs})
	_= retVal
}

func (this *Interior) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Interior) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Interior) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

