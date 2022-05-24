package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// F08B45F1-8F23-4156-9D63-1820C0ED229A
var IID_OMathBar = syscall.GUID{0xF08B45F1, 0x8F23, 0x4156, 
	[8]byte{0x9D, 0x63, 0x18, 0x20, 0xC0, 0xED, 0x22, 0x9A}}

type OMathBar struct {
	ole.OleClient
}

func NewOMathBar(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathBar {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathBar{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathBarFromVar(v ole.Variant) *OMathBar {
	return NewOMathBar(v.IDispatch(), false, false)
}

func (this *OMathBar) IID() *syscall.GUID {
	return &IID_OMathBar
}

func (this *OMathBar) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathBar) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathBar) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathBar) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathBar) E() *OMath {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathBar) BarTop() bool {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBar) SetBarTop(rhs bool)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

