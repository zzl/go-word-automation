package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 65DF9F31-B1E3-4651-87E8-51D55F302161
var IID_CoAuthoring = syscall.GUID{0x65DF9F31, 0xB1E3, 0x4651, 
	[8]byte{0x87, 0xE8, 0x51, 0xD5, 0x5F, 0x30, 0x21, 0x61}}

type CoAuthoring struct {
	ole.OleClient
}

func NewCoAuthoring(pDisp *win32.IDispatch, addRef bool, scoped bool) *CoAuthoring {
	 if pDisp == nil {
		return nil;
	}
	p := &CoAuthoring{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CoAuthoringFromVar(v ole.Variant) *CoAuthoring {
	return NewCoAuthoring(v.IDispatch(), false, false)
}

func (this *CoAuthoring) IID() *syscall.GUID {
	return &IID_CoAuthoring
}

func (this *CoAuthoring) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CoAuthoring) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CoAuthoring) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CoAuthoring) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CoAuthoring) Authors() *CoAuthors {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewCoAuthors(retVal.IDispatch(), false, true)
}

func (this *CoAuthoring) Me() *CoAuthor {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewCoAuthor(retVal.IDispatch(), false, true)
}

func (this *CoAuthoring) PendingUpdates() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CoAuthoring) Locks() *CoAuthLocks {
	retVal, _ := this.PropGet(0x00000005, nil)
	return NewCoAuthLocks(retVal.IDispatch(), false, true)
}

func (this *CoAuthoring) Updates() *CoAuthUpdates {
	retVal, _ := this.PropGet(0x00000006, nil)
	return NewCoAuthUpdates(retVal.IDispatch(), false, true)
}

func (this *CoAuthoring) Conflicts() *Conflicts {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewConflicts(retVal.IDispatch(), false, true)
}

func (this *CoAuthoring) CanShare() bool {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CoAuthoring) CanMerge() bool {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

