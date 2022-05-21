package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020925-0000-0000-C000-000000000046
var IID_DropDown = syscall.GUID{0x00020925, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DropDown struct {
	ole.OleClient
}

func NewDropDown(pDisp *win32.IDispatch, addRef bool, scoped bool) *DropDown {
	p := &DropDown{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DropDownFromVar(v ole.Variant) *DropDown {
	return NewDropDown(v.PdispValVal(), false, false)
}

func (this *DropDown) IID() *syscall.GUID {
	return &IID_DropDown
}

func (this *DropDown) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DropDown) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *DropDown) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *DropDown) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DropDown) Valid() bool {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DropDown) Default() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *DropDown) SetDefault(rhs int32)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *DropDown) Value() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *DropDown) SetValue(rhs int32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *DropDown) ListEntries() *ListEntries {
	retVal := this.PropGet(0x00000003, nil)
	return NewListEntries(retVal.PdispValVal(), false, true)
}

