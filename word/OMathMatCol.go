package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// CAE36175-3818-4C60-BCBF-0645D51EB33B
var IID_OMathMatCol = syscall.GUID{0xCAE36175, 0x3818, 0x4C60, 
	[8]byte{0xBC, 0xBF, 0x06, 0x45, 0xD5, 0x1E, 0xB3, 0x3B}}

type OMathMatCol struct {
	ole.OleClient
}

func NewOMathMatCol(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathMatCol {
	p := &OMathMatCol{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathMatColFromVar(v ole.Variant) *OMathMatCol {
	return NewOMathMatCol(v.PdispValVal(), false, false)
}

func (this *OMathMatCol) IID() *syscall.GUID {
	return &IID_OMathMatCol
}

func (this *OMathMatCol) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathMatCol) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathMatCol) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathMatCol) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathMatCol) Args() *OMathArgs {
	retVal := this.PropGet(0x00000067, nil)
	return NewOMathArgs(retVal.PdispValVal(), false, true)
}

func (this *OMathMatCol) ColIndex() int32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *OMathMatCol) Align() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *OMathMatCol) SetAlign(rhs int32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *OMathMatCol) Delete()  {
	retVal := this.Call(0x000000c8, nil)
	_= retVal
}

