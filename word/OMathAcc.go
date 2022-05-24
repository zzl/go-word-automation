package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// F258DE05-C41B-4C33-A778-F0D3F98CEEB3
var IID_OMathAcc = syscall.GUID{0xF258DE05, 0xC41B, 0x4C33, 
	[8]byte{0xA7, 0x78, 0xF0, 0xD3, 0xF9, 0x8C, 0xEE, 0xB3}}

type OMathAcc struct {
	ole.OleClient
}

func NewOMathAcc(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathAcc {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathAcc{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathAccFromVar(v ole.Variant) *OMathAcc {
	return NewOMathAcc(v.IDispatch(), false, false)
}

func (this *OMathAcc) IID() *syscall.GUID {
	return &IID_OMathAcc
}

func (this *OMathAcc) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathAcc) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathAcc) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathAcc) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathAcc) E() *OMath {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathAcc) Char() int16 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.IValVal()
}

func (this *OMathAcc) SetChar(rhs int16)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

