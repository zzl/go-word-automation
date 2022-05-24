package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0D951ADF-10A6-4C9B-BCD9-0FB8CBAD9A87
var IID_OMathFunc = syscall.GUID{0x0D951ADF, 0x10A6, 0x4C9B, 
	[8]byte{0xBC, 0xD9, 0x0F, 0xB8, 0xCB, 0xAD, 0x9A, 0x87}}

type OMathFunc struct {
	ole.OleClient
}

func NewOMathFunc(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathFunc {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathFunc{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathFuncFromVar(v ole.Variant) *OMathFunc {
	return NewOMathFunc(v.IDispatch(), false, false)
}

func (this *OMathFunc) IID() *syscall.GUID {
	return &IID_OMathFunc
}

func (this *OMathFunc) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathFunc) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathFunc) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathFunc) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathFunc) FName() *OMath {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathFunc) E() *OMath {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

