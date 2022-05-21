package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020931-0000-0000-C000-000000000046
var IID_LinkFormat = syscall.GUID{0x00020931, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type LinkFormat struct {
	ole.OleClient
}

func NewLinkFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *LinkFormat {
	p := &LinkFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LinkFormatFromVar(v ole.Variant) *LinkFormat {
	return NewLinkFormat(v.PdispValVal(), false, false)
}

func (this *LinkFormat) IID() *syscall.GUID {
	return &IID_LinkFormat
}

func (this *LinkFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *LinkFormat) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *LinkFormat) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *LinkFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *LinkFormat) AutoUpdate() bool {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *LinkFormat) SetAutoUpdate(rhs bool)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *LinkFormat) SourceName() string {
	retVal := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LinkFormat) SourcePath() string {
	retVal := this.PropGet(0x00000005, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LinkFormat) Locked() bool {
	retVal := this.PropGet(0x0000000d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *LinkFormat) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x0000000d, []interface{}{rhs})
	_= retVal
}

func (this *LinkFormat) Type() int32 {
	retVal := this.PropGet(0x00000010, nil)
	return retVal.LValVal()
}

func (this *LinkFormat) SourceFullName() string {
	retVal := this.PropGet(0x00000015, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LinkFormat) SetSourceFullName(rhs string)  {
	retVal := this.PropPut(0x00000015, []interface{}{rhs})
	_= retVal
}

func (this *LinkFormat) SavePictureWithDocument() bool {
	retVal := this.PropGet(0x00000016, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *LinkFormat) SetSavePictureWithDocument(rhs bool)  {
	retVal := this.PropPut(0x00000016, []interface{}{rhs})
	_= retVal
}

func (this *LinkFormat) BreakLink()  {
	retVal := this.Call(0x00000068, nil)
	_= retVal
}

func (this *LinkFormat) Update()  {
	retVal := this.Call(0x00000069, nil)
	_= retVal
}

