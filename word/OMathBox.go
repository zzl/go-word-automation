package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 842C37FE-C76F-4B2B-9B60-C408CB5E838E
var IID_OMathBox = syscall.GUID{0x842C37FE, 0xC76F, 0x4B2B, 
	[8]byte{0x9B, 0x60, 0xC4, 0x08, 0xCB, 0x5E, 0x83, 0x8E}}

type OMathBox struct {
	ole.OleClient
}

func NewOMathBox(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathBox {
	p := &OMathBox{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathBoxFromVar(v ole.Variant) *OMathBox {
	return NewOMathBox(v.PdispValVal(), false, false)
}

func (this *OMathBox) IID() *syscall.GUID {
	return &IID_OMathBox
}

func (this *OMathBox) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathBox) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathBox) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathBox) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathBox) E() *OMath {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathBox) OpEmu() bool {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBox) SetOpEmu(rhs bool)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *OMathBox) NoBreak() bool {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBox) SetNoBreak(rhs bool)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *OMathBox) Diff() bool {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathBox) SetDiff(rhs bool)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

