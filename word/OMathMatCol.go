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
	 if pDisp == nil {
		return nil;
	}
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
	return NewOMathMatCol(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathMatCol) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathMatCol) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathMatCol) Args() *OMathArgs {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewOMathArgs(retVal.IDispatch(), false, true)
}

func (this *OMathMatCol) ColIndex() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *OMathMatCol) Align() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *OMathMatCol) SetAlign(rhs int32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *OMathMatCol) Delete()  {
	retVal, _ := this.Call(0x000000c8, nil)
	_= retVal
}

