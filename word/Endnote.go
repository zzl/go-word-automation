package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002093E-0000-0000-C000-000000000046
var IID_Endnote = syscall.GUID{0x0002093E, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Endnote struct {
	ole.OleClient
}

func NewEndnote(pDisp *win32.IDispatch, addRef bool, scoped bool) *Endnote {
	p := &Endnote{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EndnoteFromVar(v ole.Variant) *Endnote {
	return NewEndnote(v.PdispValVal(), false, false)
}

func (this *Endnote) IID() *syscall.GUID {
	return &IID_Endnote
}

func (this *Endnote) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Endnote) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Endnote) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Endnote) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Endnote) Range() *Range {
	retVal := this.PropGet(0x00000004, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Endnote) Reference() *Range {
	retVal := this.PropGet(0x00000005, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Endnote) Index() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Endnote) Delete()  {
	retVal := this.Call(0x0000000a, nil)
	_= retVal
}

