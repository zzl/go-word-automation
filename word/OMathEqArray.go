package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 1F998A61-71C6-44C2-A0F2-1D66169B47CB
var IID_OMathEqArray = syscall.GUID{0x1F998A61, 0x71C6, 0x44C2, 
	[8]byte{0xA0, 0xF2, 0x1D, 0x66, 0x16, 0x9B, 0x47, 0xCB}}

type OMathEqArray struct {
	ole.OleClient
}

func NewOMathEqArray(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathEqArray {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathEqArray{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathEqArrayFromVar(v ole.Variant) *OMathEqArray {
	return NewOMathEqArray(v.IDispatch(), false, false)
}

func (this *OMathEqArray) IID() *syscall.GUID {
	return &IID_OMathEqArray
}

func (this *OMathEqArray) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathEqArray) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathEqArray) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathEqArray) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathEqArray) E() *OMathArgs {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMathArgs(retVal.IDispatch(), false, true)
}

func (this *OMathEqArray) MaxDist() bool {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathEqArray) SetMaxDist(rhs bool)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *OMathEqArray) ObjDist() bool {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathEqArray) SetObjDist(rhs bool)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *OMathEqArray) Align() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *OMathEqArray) SetAlign(rhs int32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *OMathEqArray) RowSpacingRule() int32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.LValVal()
}

func (this *OMathEqArray) SetRowSpacingRule(rhs int32)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *OMathEqArray) RowSpacing() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *OMathEqArray) SetRowSpacing(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

