package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// BF043168-F4DE-4E7C-B206-741A8B3EF71A
var IID_EndnoteOptions = syscall.GUID{0xBF043168, 0xF4DE, 0x4E7C, 
	[8]byte{0xB2, 0x06, 0x74, 0x1A, 0x8B, 0x3E, 0xF7, 0x1A}}

type EndnoteOptions struct {
	ole.OleClient
}

func NewEndnoteOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *EndnoteOptions {
	p := &EndnoteOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EndnoteOptionsFromVar(v ole.Variant) *EndnoteOptions {
	return NewEndnoteOptions(v.PdispValVal(), false, false)
}

func (this *EndnoteOptions) IID() *syscall.GUID {
	return &IID_EndnoteOptions
}

func (this *EndnoteOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *EndnoteOptions) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *EndnoteOptions) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *EndnoteOptions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *EndnoteOptions) Location() int32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *EndnoteOptions) SetLocation(rhs int32)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *EndnoteOptions) NumberStyle() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *EndnoteOptions) SetNumberStyle(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *EndnoteOptions) StartingNumber() int32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *EndnoteOptions) SetStartingNumber(rhs int32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *EndnoteOptions) NumberingRule() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *EndnoteOptions) SetNumberingRule(rhs int32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

