package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 5DAA8BB6-054E-48F6-BEAC-EAAD02BE0CC7
var IID_OMathMatRow = syscall.GUID{0x5DAA8BB6, 0x054E, 0x48F6, 
	[8]byte{0xBE, 0xAC, 0xEA, 0xAD, 0x02, 0xBE, 0x0C, 0xC7}}

type OMathMatRow struct {
	ole.OleClient
}

func NewOMathMatRow(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathMatRow {
	p := &OMathMatRow{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathMatRowFromVar(v ole.Variant) *OMathMatRow {
	return NewOMathMatRow(v.PdispValVal(), false, false)
}

func (this *OMathMatRow) IID() *syscall.GUID {
	return &IID_OMathMatRow
}

func (this *OMathMatRow) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathMatRow) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathMatRow) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathMatRow) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathMatRow) Args() *OMathArgs {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMathArgs(retVal.PdispValVal(), false, true)
}

func (this *OMathMatRow) RowIndex() int32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *OMathMatRow) Delete()  {
	retVal := this.Call(0x000000c8, nil)
	_= retVal
}

