package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 194F8476-B79D-4572-A609-294207DE77C1
var IID_ErrorBars = syscall.GUID{0x194F8476, 0xB79D, 0x4572, 
	[8]byte{0xA6, 0x09, 0x29, 0x42, 0x07, 0xDE, 0x77, 0xC1}}

type ErrorBars struct {
	ole.OleClient
}

func NewErrorBars(pDisp *win32.IDispatch, addRef bool, scoped bool) *ErrorBars {
	 if pDisp == nil {
		return nil;
	}
	p := &ErrorBars{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ErrorBarsFromVar(v ole.Variant) *ErrorBars {
	return NewErrorBars(v.IDispatch(), false, false)
}

func (this *ErrorBars) IID() *syscall.GUID {
	return &IID_ErrorBars
}

func (this *ErrorBars) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ErrorBars) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ErrorBars) Name() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ErrorBars) Select() ole.Variant {
	retVal, _ := this.Call(0x000000eb, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ErrorBars) Border() *ChartBorder {
	retVal, _ := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.IDispatch(), false, true)
}

func (this *ErrorBars) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ErrorBars) ClearFormats() ole.Variant {
	retVal, _ := this.Call(0x00000070, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ErrorBars) EndStyle() int32 {
	retVal, _ := this.PropGet(0x00000464, nil)
	return retVal.LValVal()
}

func (this *ErrorBars) SetEndStyle(rhs int32)  {
	_ = this.PropPut(0x00000464, []interface{}{rhs})
}

func (this *ErrorBars) Format() *ChartFormat {
	retVal, _ := this.PropGet(0x60020008, nil)
	return NewChartFormat(retVal.IDispatch(), false, true)
}

func (this *ErrorBars) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ErrorBars) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

