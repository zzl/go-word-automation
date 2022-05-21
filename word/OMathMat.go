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
	return NewOMathMat(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathMat) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathMat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathMat) Rows() *OMathMatRows {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMathMatRows(retVal.PdispValVal(), false, true)
}

func (this *OMathMat) Cols() *OMathMatCols {
	retVal := this.PropGet(0x00000068, nil)
	return NewOMathMatCols(retVal.PdispValVal(), false, true)
}

func (this *OMathMat) Cell(row int32, col int32) *OMath {
	retVal := this.PropGet(0x00000069, []interface{}{row, col})
	return NewOMath(retVal.PdispValVal(), false, true)
}

func (this *OMathMat) Align() int32 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetAlign(rhs int32)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *OMathMat) PlcHoldHidden() bool {
	retVal := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathMat) SetPlcHoldHidden(rhs bool)  {
	retVal := this.PropPut(0x0000006b, []interface{}{rhs})
	_= retVal
}

func (this *OMathMat) RowSpacingRule() int32 {
	retVal := this.PropGet(0x0000006c, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetRowSpacingRule(rhs int32)  {
	retVal := this.PropPut(0x0000006c, []interface{}{rhs})
	_= retVal
}

func (this *OMathMat) RowSpacing() int32 {
	retVal := this.PropGet(0x0000006d, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetRowSpacing(rhs int32)  {
	retVal := this.PropPut(0x0000006d, []interface{}{rhs})
	_= retVal
}

func (this *OMathMat) ColSpacing() int32 {
	retVal := this.PropGet(0x0000006e, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetColSpacing(rhs int32)  {
	retVal := this.PropPut(0x0000006e, []interface{}{rhs})
	_= retVal
}

func (this *OMathMat) ColGapRule() int32 {
	retVal := this.PropGet(0x0000006f, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetColGapRule(rhs int32)  {
	retVal := this.PropPut(0x0000006f, []interface{}{rhs})
	_= retVal
}

func (this *OMathMat) ColGap() int32 {
	retVal := this.PropGet(0x00000070, nil)
	return retVal.LValVal()
}

func (this *OMathMat) SetColGap(rhs int32)  {
	retVal := this.PropPut(0x00000070, []interface{}{rhs})
	_= retVal
}

