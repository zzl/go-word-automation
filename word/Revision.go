package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"time"
)

// 00020981-0000-0000-C000-000000000046
var IID_Revision = syscall.GUID{0x00020981, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Revision struct {
	ole.OleClient
}

func NewRevision(pDisp *win32.IDispatch, addRef bool, scoped bool) *Revision {
	 if pDisp == nil {
		return nil;
	}
	p := &Revision{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RevisionFromVar(v ole.Variant) *Revision {
	return NewRevision(v.IDispatch(), false, false)
}

func (this *Revision) IID() *syscall.GUID {
	return &IID_Revision
}

func (this *Revision) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Revision) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Revision) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Revision) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Revision) Author() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Revision) Date() time.Time {
	retVal, _ := this.PropGet(0x00000002, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *Revision) Range() *Range {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Revision) Type() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Revision) Index() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Revision) Accept()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *Revision) Reject()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *Revision) Style() *Style {
	retVal, _ := this.PropGet(0x00000008, nil)
	return NewStyle(retVal.IDispatch(), false, true)
}

func (this *Revision) FormatDescription() string {
	retVal, _ := this.PropGet(0x00000009, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Revision) MovedRange() *Range {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Revision) Cells() *Cells {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return NewCells(retVal.IDispatch(), false, true)
}

