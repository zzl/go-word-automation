package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 47CEF4AE-DC32-4220-8AA5-19CCC0E6633A
var IID_Reviewer = syscall.GUID{0x47CEF4AE, 0xDC32, 0x4220, 
	[8]byte{0x8A, 0xA5, 0x19, 0xCC, 0xC0, 0xE6, 0x63, 0x3A}}

type Reviewer struct {
	ole.OleClient
}

func NewReviewer(pDisp *win32.IDispatch, addRef bool, scoped bool) *Reviewer {
	p := &Reviewer{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ReviewerFromVar(v ole.Variant) *Reviewer {
	return NewReviewer(v.PdispValVal(), false, false)
}

func (this *Reviewer) IID() *syscall.GUID {
	return &IID_Reviewer
}

func (this *Reviewer) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Reviewer) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Reviewer) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Reviewer) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Reviewer) Visible() bool {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Reviewer) SetVisible(rhs bool)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

