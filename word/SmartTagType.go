package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 5E9A888C-E5DC-4DCB-8308-3C91FB61E6F4
var IID_SmartTagType = syscall.GUID{0x5E9A888C, 0xE5DC, 0x4DCB, 
	[8]byte{0x83, 0x08, 0x3C, 0x91, 0xFB, 0x61, 0xE6, 0xF4}}

type SmartTagType struct {
	ole.OleClient
}

func NewSmartTagType(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTagType {
	p := &SmartTagType{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagTypeFromVar(v ole.Variant) *SmartTagType {
	return NewSmartTagType(v.PdispValVal(), false, false)
}

func (this *SmartTagType) IID() *syscall.GUID {
	return &IID_SmartTagType
}

func (this *SmartTagType) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTagType) Name() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTagType) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SmartTagType) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SmartTagType) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SmartTagType) SmartTagActions() *SmartTagActions {
	retVal := this.PropGet(0x000003eb, nil)
	return NewSmartTagActions(retVal.PdispValVal(), false, true)
}

func (this *SmartTagType) SmartTagRecognizers() *SmartTagRecognizers {
	retVal := this.PropGet(0x000003ec, nil)
	return NewSmartTagRecognizers(retVal.PdispValVal(), false, true)
}

func (this *SmartTagType) FriendlyName() string {
	retVal := this.PropGet(0x000003ed, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

