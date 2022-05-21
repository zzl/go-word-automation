package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 15EBE471-0182-4CCE-98D0-B6614D1C32A1
var IID_SmartTagRecognizer = syscall.GUID{0x15EBE471, 0x0182, 0x4CCE, 
	[8]byte{0x98, 0xD0, 0xB6, 0x61, 0x4D, 0x1C, 0x32, 0xA1}}

type SmartTagRecognizer struct {
	ole.OleClient
}

func NewSmartTagRecognizer(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTagRecognizer {
	p := &SmartTagRecognizer{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagRecognizerFromVar(v ole.Variant) *SmartTagRecognizer {
	return NewSmartTagRecognizer(v.PdispValVal(), false, false)
}

func (this *SmartTagRecognizer) IID() *syscall.GUID {
	return &IID_SmartTagRecognizer
}

func (this *SmartTagRecognizer) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTagRecognizer) FullName() string {
	retVal := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTagRecognizer) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SmartTagRecognizer) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SmartTagRecognizer) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SmartTagRecognizer) Enabled() bool {
	retVal := this.PropGet(0x000003eb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagRecognizer) SetEnabled(rhs bool)  {
	retVal := this.PropPut(0x000003eb, []interface{}{rhs})
	_= retVal
}

func (this *SmartTagRecognizer) ProgID() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTagRecognizer) Caption() string {
	retVal := this.PropGet(0x000003ec, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

