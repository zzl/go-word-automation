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
	 if pDisp == nil {
		return nil;
	}
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
	return NewInterior(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x60020000, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Interior) SetColor(rhs interface{})  {
	_ = this.PropPut(0x60020000, []interface{}{rhs})
}

func (this *Interior) ColorIndex() ole.Variant {
	retVal, _ := this.PropGet(0x60020002, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Interior) SetColorIndex(rhs interface{})  {
	_ = this.PropPut(0x60020002, []interface{}{rhs})
}

func (this *Interior) InvertIfNegative() ole.Variant {
	retVal, _ := this.PropGet(0x60020004, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Interior) SetInvertIfNegative(rhs interface{})  {
	_ = this.PropPut(0x60020004, []interface{}{rhs})
}

func (this *Interior) Pattern() ole.Variant {
	retVal, _ := this.PropGet(0x60020006, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Interior) SetPattern(rhs interface{})  {
	_ = this.PropPut(0x60020006, []interface{}{rhs})
}

func (this *Interior) PatternColor() ole.Variant {
	retVal, _ := this.PropGet(0x60020008, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Interior) SetPatternColor(rhs interface{})  {
	_ = this.PropPut(0x60020008, []interface{}{rhs})
}

func (this *Interior) PatternColorIndex() ole.Variant {
	retVal, _ := this.PropGet(0x6002000a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Interior) SetPatternColorIndex(rhs interface{})  {
	_ = this.PropPut(0x6002000a, []interface{}{rhs})
}

func (this *Interior) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Interior) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *Interior) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

