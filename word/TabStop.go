package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020954-0000-0000-C000-000000000046
var IID_TabStop = syscall.GUID{0x00020954, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TabStop struct {
	ole.OleClient
}

func NewTabStop(pDisp *win32.IDispatch, addRef bool, scoped bool) *TabStop {
	p := &TabStop{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TabStopFromVar(v ole.Variant) *TabStop {
	return NewTabStop(v.PdispValVal(), false, false)
}

func (this *TabStop) IID() *syscall.GUID {
	return &IID_TabStop
}

func (this *TabStop) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TabStop) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *TabStop) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TabStop) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *TabStop) Alignment() int32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *TabStop) SetAlignment(rhs int32)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *TabStop) Leader() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *TabStop) SetLeader(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *TabStop) Position() float32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *TabStop) SetPosition(rhs float32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *TabStop) CustomTab() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TabStop) Next() *TabStop {
	retVal := this.PropGet(0x00000068, nil)
	return NewTabStop(retVal.PdispValVal(), false, true)
}

func (this *TabStop) Previous() *TabStop {
	retVal := this.PropGet(0x00000069, nil)
	return NewTabStop(retVal.PdispValVal(), false, true)
}

func (this *TabStop) Clear()  {
	retVal := this.Call(0x000000c8, nil)
	_= retVal
}

