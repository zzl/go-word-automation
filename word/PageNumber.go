package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020987-0000-0000-C000-000000000046
var IID_PageNumber = syscall.GUID{0x00020987, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type PageNumber struct {
	ole.OleClient
}

func NewPageNumber(pDisp *win32.IDispatch, addRef bool, scoped bool) *PageNumber {
	 if pDisp == nil {
		return nil;
	}
	p := &PageNumber{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PageNumberFromVar(v ole.Variant) *PageNumber {
	return NewPageNumber(v.IDispatch(), false, false)
}

func (this *PageNumber) IID() *syscall.GUID {
	return &IID_PageNumber
}

func (this *PageNumber) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *PageNumber) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *PageNumber) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *PageNumber) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *PageNumber) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *PageNumber) Alignment() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *PageNumber) SetAlignment(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *PageNumber) Select()  {
	retVal, _ := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *PageNumber) Copy()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *PageNumber) Cut()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *PageNumber) Delete()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

