package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020969-0000-0000-C000-000000000046
var IID_RoutingSlip = syscall.GUID{0x00020969, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type RoutingSlip struct {
	ole.OleClient
}

func NewRoutingSlip(pDisp *win32.IDispatch, addRef bool, scoped bool) *RoutingSlip {
	p := &RoutingSlip{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RoutingSlipFromVar(v ole.Variant) *RoutingSlip {
	return NewRoutingSlip(v.PdispValVal(), false, false)
}

func (this *RoutingSlip) IID() *syscall.GUID {
	return &IID_RoutingSlip
}

func (this *RoutingSlip) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *RoutingSlip) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *RoutingSlip) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *RoutingSlip) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *RoutingSlip) Subject() string {
	retVal := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *RoutingSlip) SetSubject(rhs string)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *RoutingSlip) Message() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *RoutingSlip) SetMessage(rhs string)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *RoutingSlip) Delivery() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *RoutingSlip) SetDelivery(rhs int32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *RoutingSlip) TrackStatus() bool {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *RoutingSlip) SetTrackStatus(rhs bool)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *RoutingSlip) Protect() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *RoutingSlip) SetProtect(rhs int32)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *RoutingSlip) ReturnWhenDone() bool {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *RoutingSlip) SetReturnWhenDone(rhs bool)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *RoutingSlip) Status() int32 {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

var RoutingSlip_Recipients_OptArgs= []string{
	"Index", 
}

func (this *RoutingSlip) Recipients(optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(RoutingSlip_Recipients_OptArgs, optArgs)
	retVal := this.PropGet(0x00000009, nil, optArgs...)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *RoutingSlip) Reset()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *RoutingSlip) AddRecipient(recipient string)  {
	retVal := this.Call(0x00000066, []interface{}{recipient})
	_= retVal
}

