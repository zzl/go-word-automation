package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002093F-0000-0000-C000-000000000046
var IID_Footnote = syscall.GUID{0x0002093F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Footnote struct {
	ole.OleClient
}

func NewFootnote(pDisp *win32.IDispatch, addRef bool, scoped bool) *Footnote {
	p := &Footnote{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FootnoteFromVar(v ole.Variant) *Footnote {
	return NewFootnote(v.PdispValVal(), false, false)
}

func (this *Footnote) IID() *syscall.GUID {
	return &IID_Footnote
}

func (this *Footnote) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Footnote) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Footnote) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Footnote) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Footnote) Range() *Range {
	retVal := this.PropGet(0x00000004, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Footnote) Reference() *Range {
	retVal := this.PropGet(0x00000005, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Footnote) Index() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Footnote) Delete()  {
	retVal := this.Call(0x0000000a, nil)
	_= retVal
}

