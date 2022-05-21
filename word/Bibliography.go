package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 3834F60F-EE8C-455D-A441-D766675D6D3B
var IID_Bibliography = syscall.GUID{0x3834F60F, 0xEE8C, 0x455D, 
	[8]byte{0xA4, 0x41, 0xD7, 0x66, 0x67, 0x5D, 0x6D, 0x3B}}

type Bibliography struct {
	ole.OleClient
}

func NewBibliography(pDisp *win32.IDispatch, addRef bool, scoped bool) *Bibliography {
	p := &Bibliography{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BibliographyFromVar(v ole.Variant) *Bibliography {
	return NewBibliography(v.PdispValVal(), false, false)
}

func (this *Bibliography) IID() *syscall.GUID {
	return &IID_Bibliography
}

func (this *Bibliography) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Bibliography) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Bibliography) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Bibliography) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Bibliography) Sources() *Sources {
	retVal := this.PropGet(0x00000067, nil)
	return NewSources(retVal.PdispValVal(), false, true)
}

func (this *Bibliography) BibliographyStyle() string {
	retVal := this.PropGet(0x00000069, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Bibliography) SetBibliographyStyle(rhs string)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *Bibliography) GenerateUniqueTag() string {
	retVal := this.Call(0x00000068, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

