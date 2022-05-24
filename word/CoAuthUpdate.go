package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 9E6B5EC5-E8E4-40AF-9540-6203F71E2823
var IID_CoAuthUpdate = syscall.GUID{0x9E6B5EC5, 0xE8E4, 0x40AF, 
	[8]byte{0x95, 0x40, 0x62, 0x03, 0xF7, 0x1E, 0x28, 0x23}}

type CoAuthUpdate struct {
	ole.OleClient
}

func NewCoAuthUpdate(pDisp *win32.IDispatch, addRef bool, scoped bool) *CoAuthUpdate {
	 if pDisp == nil {
		return nil;
	}
	p := &CoAuthUpdate{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CoAuthUpdateFromVar(v ole.Variant) *CoAuthUpdate {
	return NewCoAuthUpdate(v.IDispatch(), false, false)
}

func (this *CoAuthUpdate) IID() *syscall.GUID {
	return &IID_CoAuthUpdate
}

func (this *CoAuthUpdate) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CoAuthUpdate) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CoAuthUpdate) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CoAuthUpdate) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CoAuthUpdate) Range() *Range {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

