package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 99755F80-FE96-4F7D-B636-B8E800E54F44
var IID_CoAuthLock = syscall.GUID{0x99755F80, 0xFE96, 0x4F7D, 
	[8]byte{0xB6, 0x36, 0xB8, 0xE8, 0x00, 0xE5, 0x4F, 0x44}}

type CoAuthLock struct {
	ole.OleClient
}

func NewCoAuthLock(pDisp *win32.IDispatch, addRef bool, scoped bool) *CoAuthLock {
	p := &CoAuthLock{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CoAuthLockFromVar(v ole.Variant) *CoAuthLock {
	return NewCoAuthLock(v.PdispValVal(), false, false)
}

func (this *CoAuthLock) IID() *syscall.GUID {
	return &IID_CoAuthLock
}

func (this *CoAuthLock) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CoAuthLock) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CoAuthLock) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CoAuthLock) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CoAuthLock) Type() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *CoAuthLock) Owner() *CoAuthor {
	retVal := this.PropGet(0x00000002, nil)
	return NewCoAuthor(retVal.PdispValVal(), false, true)
}

func (this *CoAuthLock) Range() *Range {
	retVal := this.PropGet(0x00000003, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *CoAuthLock) HeaderFooter() bool {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CoAuthLock) Unlock()  {
	retVal := this.Call(0x00000006, nil)
	_= retVal
}

