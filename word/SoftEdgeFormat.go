package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// D040DAF9-6CE4-4BE8-839D-F4538A4301CF
var IID_SoftEdgeFormat = syscall.GUID{0xD040DAF9, 0x6CE4, 0x4BE8, 
	[8]byte{0x83, 0x9D, 0xF4, 0x53, 0x8A, 0x43, 0x01, 0xCF}}

type SoftEdgeFormat struct {
	ole.OleClient
}

func NewSoftEdgeFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *SoftEdgeFormat {
	p := &SoftEdgeFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SoftEdgeFormatFromVar(v ole.Variant) *SoftEdgeFormat {
	return NewSoftEdgeFormat(v.PdispValVal(), false, false)
}

func (this *SoftEdgeFormat) IID() *syscall.GUID {
	return &IID_SoftEdgeFormat
}

func (this *SoftEdgeFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SoftEdgeFormat) Type() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *SoftEdgeFormat) SetType(rhs int32)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *SoftEdgeFormat) Radius() float32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.FltValVal()
}

func (this *SoftEdgeFormat) SetRadius(rhs float32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *SoftEdgeFormat) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SoftEdgeFormat) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SoftEdgeFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

