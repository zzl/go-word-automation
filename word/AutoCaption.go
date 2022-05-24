package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002097B-0000-0000-C000-000000000046
var IID_AutoCaption = syscall.GUID{0x0002097B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AutoCaption struct {
	ole.OleClient
}

func NewAutoCaption(pDisp *win32.IDispatch, addRef bool, scoped bool) *AutoCaption {
	 if pDisp == nil {
		return nil;
	}
	p := &AutoCaption{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AutoCaptionFromVar(v ole.Variant) *AutoCaption {
	return NewAutoCaption(v.IDispatch(), false, false)
}

func (this *AutoCaption) IID() *syscall.GUID {
	return &IID_AutoCaption
}

func (this *AutoCaption) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AutoCaption) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AutoCaption) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AutoCaption) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AutoCaption) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *AutoCaption) AutoInsert() bool {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCaption) SetAutoInsert(rhs bool)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *AutoCaption) Index() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *AutoCaption) CaptionLabel() ole.Variant {
	retVal, _ := this.PropGet(0x00000003, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *AutoCaption) SetCaptionLabel(rhs *ole.Variant)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

