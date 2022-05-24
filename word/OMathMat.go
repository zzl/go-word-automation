package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 3E061A7E-67AD-4EAA-BC1E-55057D5E596F
var IID_OMathMat = syscall.GUID{0x3E061A7E, 0x67AD, 0x4EAA, 
	[8]byte{0xBC, 0x1E, 0x55, 0x05, 0x7D, 0x5E, 0x59, 0x6F}}

type OMathMat struct {
	ole.OleClient
}

func NewOMathMat(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathMat {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathMat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathMatFromVar(v ole.Variant) *OMathMat {
	return NewOMathMat(v.IDispatch(), false, false)
}

func (this *OMathMat) IID() *syscall.GUID {
	return &IID_OMathMat
}

func (this *OMathMat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathMat) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathMat) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathMat) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathMat) Rows() *OMathMatRows {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMathMatRows(retVal.IDispatch(), false, true)
}

func (this *OMathMat) Cols() *OMathMatCols {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewOMathMatCols(retVal.IDispatch(), false, true)
}

func (this *OMathMat) Cell(row int32, col int32) *OMath {
	retVal, _ := this.PropGet(0x00000069, []interface{}{row, col})
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathMat) Align() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetAlign(rhs int32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *OMathMat) PlcHoldHidden() bool {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathMat) SetPlcHoldHidden(rhs bool)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *OMathMat) RowSpacingRule() int32 {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetRowSpacingRule(rhs int32)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *OMathMat) RowSpacing() int32 {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetRowSpacing(rhs int32)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *OMathMat) ColSpacing() int32 {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetColSpacing(rhs int32)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *OMathMat) ColGapRule() int32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetColGapRule(rhs int32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *OMathMat) ColGap() int32 {
	retVal, _ := this.PropGet(0x00000070, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetColGap(rhs int32)  {
	_ = this.PropPut(0x00000070, []interface{}{rhs})
}

