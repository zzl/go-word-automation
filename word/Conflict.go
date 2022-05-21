package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 6215E4B1-545A-406E-9824-0A5B5AC8AD21
var IID_Conflict = syscall.GUID{0x6215E4B1, 0x545A, 0x406E, 
	[8]byte{0x98, 0x24, 0x0A, 0x5B, 0x5A, 0xC8, 0xAD, 0x21}}

type Conflict struct {
	ole.OleClient
}

func NewConflict(pDisp *win32.IDispatch, addRef bool, scoped bool) *Conflict {
	p := &Conflict{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ConflictFromVar(v ole.Variant) *Conflict {
	return NewConflict(v.PdispValVal(), false, false)
}

func (this *Conflict) IID() *syscall.GUID {
	return &IID_Conflict
}

func (this *Conflict) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Conflict) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Conflict) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Conflict) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Conflict) Range() *Range {
	retVal := this.PropGet(0x00000003, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Conflict) Type() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Conflict) Index() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Conflict) Accept()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *Conflict) Reject()  {
	retVal := this.Call(0x00000066, nil)
	_= retVal
}

