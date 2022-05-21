package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020933-0000-0000-C000-000000000046
var IID_OLEFormat = syscall.GUID{0x00020933, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEFormat struct {
	ole.OleClient
}

func NewOLEFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *OLEFormat {
	p := &OLEFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OLEFormatFromVar(v ole.Variant) *OLEFormat {
	return NewOLEFormat(v.PdispValVal(), false, false)
}

func (this *OLEFormat) IID() *syscall.GUID {
	return &IID_OLEFormat
}

func (this *OLEFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OLEFormat) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OLEFormat) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *OLEFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OLEFormat) ClassType() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEFormat) SetClassType(rhs string)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *OLEFormat) DisplayAsIcon() bool {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEFormat) SetDisplayAsIcon(rhs bool)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *OLEFormat) IconName() string {
	retVal := this.PropGet(0x00000007, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEFormat) SetIconName(rhs string)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *OLEFormat) IconPath() string {
	retVal := this.PropGet(0x00000008, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEFormat) IconIndex() int32 {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.LValVal()
}

func (this *OLEFormat) SetIconIndex(rhs int32)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *OLEFormat) IconLabel() string {
	retVal := this.PropGet(0x0000000a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEFormat) SetIconLabel(rhs string)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *OLEFormat) Label() string {
	retVal := this.PropGet(0x0000000c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEFormat) Object() *ole.DispatchClass {
	retVal := this.PropGet(0x0000000e, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OLEFormat) ProgID() string {
	retVal := this.PropGet(0x00000016, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEFormat) Activate()  {
	retVal := this.Call(0x00000068, nil)
	_= retVal
}

func (this *OLEFormat) Edit()  {
	retVal := this.Call(0x0000006a, nil)
	_= retVal
}

func (this *OLEFormat) Open()  {
	retVal := this.Call(0x0000006b, nil)
	_= retVal
}

var OLEFormat_DoVerb_OptArgs= []string{
	"VerbIndex", 
}

func (this *OLEFormat) DoVerb(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(OLEFormat_DoVerb_OptArgs, optArgs)
	retVal := this.Call(0x0000006d, nil, optArgs...)
	_= retVal
}

var OLEFormat_ConvertTo_OptArgs= []string{
	"ClassType", "DisplayAsIcon", "IconFileName", "IconIndex", "IconLabel", 
}

func (this *OLEFormat) ConvertTo(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(OLEFormat_ConvertTo_OptArgs, optArgs)
	retVal := this.Call(0x0000006e, nil, optArgs...)
	_= retVal
}

func (this *OLEFormat) ActivateAs(classType string)  {
	retVal := this.Call(0x0000006f, []interface{}{classType})
	_= retVal
}

func (this *OLEFormat) PreserveFormattingOnUpdate() bool {
	retVal := this.PropGet(0x00000070, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OLEFormat) SetPreserveFormattingOnUpdate(rhs bool)  {
	retVal := this.PropPut(0x00000070, []interface{}{rhs})
	_= retVal
}

