package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020935-0000-0000-C000-000000000046
var IID_System = syscall.GUID{0x00020935, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type System struct {
	ole.OleClient
}

func NewSystem(pDisp *win32.IDispatch, addRef bool, scoped bool) *System {
	p := &System{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SystemFromVar(v ole.Variant) *System {
	return NewSystem(v.PdispValVal(), false, false)
}

func (this *System) IID() *syscall.GUID {
	return &IID_System
}

func (this *System) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *System) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *System) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *System) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *System) OperatingSystem() string {
	retVal := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *System) ProcessorType() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *System) Version() string {
	retVal := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *System) FreeDiskSpace() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *System) Country() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *System) LanguageDesignation() string {
	retVal := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *System) HorizontalResolution() int32 {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *System) VerticalResolution() int32 {
	retVal := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *System) ProfileString(section string, key string) string {
	retVal := this.PropGet(0x00000009, []interface{}{section, key})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *System) SetProfileString(section string, key string, rhs string)  {
	retVal := this.PropPut(0x00000009, []interface{}{section, key, rhs})
	_= retVal
}

func (this *System) PrivateProfileString(fileName string, section string, key string) string {
	retVal := this.PropGet(0x0000000a, []interface{}{fileName, section, key})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *System) SetPrivateProfileString(fileName string, section string, key string, rhs string)  {
	retVal := this.PropPut(0x0000000a, []interface{}{fileName, section, key, rhs})
	_= retVal
}

func (this *System) MathCoprocessorInstalled() bool {
	retVal := this.PropGet(0x0000000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *System) ComputerType() string {
	retVal := this.PropGet(0x0000000c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *System) MacintoshName() string {
	retVal := this.PropGet(0x0000000e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *System) QuickDrawInstalled() bool {
	retVal := this.PropGet(0x0000000f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *System) Cursor() int32 {
	retVal := this.PropGet(0x00000010, nil)
	return retVal.LValVal()
}

func (this *System) SetCursor(rhs int32)  {
	retVal := this.PropPut(0x00000010, []interface{}{rhs})
	_= retVal
}

func (this *System) MSInfo()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

var System_Connect_OptArgs= []string{
	"Drive", "Password", 
}

func (this *System) Connect(path string, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(System_Connect_OptArgs, optArgs)
	retVal := this.Call(0x00000066, []interface{}{path}, optArgs...)
	_= retVal
}

func (this *System) CountryRegion() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

