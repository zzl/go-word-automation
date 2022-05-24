package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// B923FDE0-F08C-11D3-91B0-00105A0A19FD
var IID_CustomProperty = syscall.GUID{0xB923FDE0, 0xF08C, 0x11D3, 
	[8]byte{0x91, 0xB0, 0x00, 0x10, 0x5A, 0x0A, 0x19, 0xFD}}

type CustomProperty struct {
	ole.OleClient
}

func NewCustomProperty(pDisp *win32.IDispatch, addRef bool, scoped bool) *CustomProperty {
	 if pDisp == nil {
		return nil;
	}
	p := &CustomProperty{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CustomPropertyFromVar(v ole.Variant) *CustomProperty {
	return NewCustomProperty(v.IDispatch(), false, false)
}

func (this *CustomProperty) IID() *syscall.GUID {
	return &IID_CustomProperty
}

func (this *CustomProperty) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CustomProperty) Name() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CustomProperty) Value() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CustomProperty) SetValue(rhs string)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *CustomProperty) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CustomProperty) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CustomProperty) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CustomProperty) Delete()  {
	retVal, _ := this.Call(0x0000000b, nil)
	_= retVal
}

