package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 50209974-BA32-4A03-8FA6-BAC56CC056FD
var IID_OMathFrac = syscall.GUID{0x50209974, 0xBA32, 0x4A03, 
	[8]byte{0x8F, 0xA6, 0xBA, 0xC5, 0x6C, 0xC0, 0x56, 0xFD}}

type OMathFrac struct {
	ole.OleClient
}

func NewOMathFrac(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathFrac {
	p := &OMathFrac{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathFracFromVar(v ole.Variant) *OMathFrac {
	return NewOMathFrac(v.PdispValVal(), false, false)
}

func (this *OMathFrac) IID() *syscall.GUID {
	return &IID_OMathFrac
}

func (this *OMathFrac) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathFrac) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathFrac) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathFrac) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathFrac) Num() *OMath {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathFrac) Den() *OMath {
	retVal := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathFrac) Type() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *OMathFrac) SetType(rhs int32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

