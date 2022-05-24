package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020966-0000-0000-C000-000000000046
var IID_Variable = syscall.GUID{0x00020966, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Variable struct {
	ole.OleClient
}

func NewVariable(pDisp *win32.IDispatch, addRef bool, scoped bool) *Variable {
	 if pDisp == nil {
		return nil;
	}
	p := &Variable{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func VariableFromVar(v ole.Variant) *Variable {
	return NewVariable(v.IDispatch(), false, false)
}

func (this *Variable) IID() *syscall.GUID {
	return &IID_Variable
}

func (this *Variable) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Variable) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Variable) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Variable) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Variable) Name() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Variable) Value() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Variable) SetValue(rhs string)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *Variable) Index() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Variable) Delete()  {
	retVal, _ := this.Call(0x0000000b, nil)
	_= retVal
}

