package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// D0A95726-678A-4B9D-8103-1E2B86735AE7
var IID_OMathScrSup = syscall.GUID{0xD0A95726, 0x678A, 0x4B9D, 
	[8]byte{0x81, 0x03, 0x1E, 0x2B, 0x86, 0x73, 0x5A, 0xE7}}

type OMathScrSup struct {
	ole.OleClient
}

func NewOMathScrSup(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathScrSup {
	p := &OMathScrSup{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathScrSupFromVar(v ole.Variant) *OMathScrSup {
	return NewOMathScrSup(v.PdispValVal(), false, false)
}

func (this *OMathScrSup) IID() *syscall.GUID {
	return &IID_OMathScrSup
}

func (this *OMathScrSup) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathScrSup) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathScrSup) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathScrSup) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathScrSup) E() *OMath {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathScrSup) Sup() *OMath {
	retVal := this.PropGet(0x00000068, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

