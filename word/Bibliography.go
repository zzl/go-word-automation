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
	 if pDisp == nil {
		return nil;
	}
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
	return NewBibliography(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Bibliography) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Bibliography) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Bibliography) Sources() *Sources {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewSources(retVal.IDispatch(), false, true)
}

func (this *Bibliography) BibliographyStyle() string {
	retVal, _ := this.PropGet(0x00000069, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Bibliography) SetBibliographyStyle(rhs string)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *Bibliography) GenerateUniqueTag() string {
	retVal, _ := this.Call(0x00000068, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

