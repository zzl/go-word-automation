package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// DB77D541-85C3-42E8-8649-AFBD7CF87866
var IID_OMathPhantom = syscall.GUID{0xDB77D541, 0x85C3, 0x42E8, 
	[8]byte{0x86, 0x49, 0xAF, 0xBD, 0x7C, 0xF8, 0x78, 0x66}}

type OMathPhantom struct {
	ole.OleClient
}

func NewOMathPhantom(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathPhantom {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathPhantom{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathPhantomFromVar(v ole.Variant) *OMathPhantom {
	return NewOMathPhantom(v.IDispatch(), false, false)
}

func (this *OMathPhantom) IID() *syscall.GUID {
	return &IID_OMathPhantom
}

func (this *OMathPhantom) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathPhantom) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathPhantom) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathPhantom) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathPhantom) E() *OMath {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathPhantom) Show() bool {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathPhantom) SetShow(rhs bool)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *OMathPhantom) ZeroWid() bool {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathPhantom) SetZeroWid(rhs bool)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *OMathPhantom) ZeroAsc() bool {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathPhantom) SetZeroAsc(rhs bool)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *OMathPhantom) ZeroDesc() bool {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathPhantom) SetZeroDesc(rhs bool)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *OMathPhantom) Transp() bool {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathPhantom) SetTransp(rhs bool)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *OMathPhantom) Smash() bool {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathPhantom) SetSmash(rhs bool)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

