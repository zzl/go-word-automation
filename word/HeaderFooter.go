package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020985-0000-0000-C000-000000000046
var IID_HeaderFooter = syscall.GUID{0x00020985, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HeaderFooter struct {
	ole.OleClient
}

func NewHeaderFooter(pDisp *win32.IDispatch, addRef bool, scoped bool) *HeaderFooter {
	 if pDisp == nil {
		return nil;
	}
	p := &HeaderFooter{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HeaderFooterFromVar(v ole.Variant) *HeaderFooter {
	return NewHeaderFooter(v.IDispatch(), false, false)
}

func (this *HeaderFooter) IID() *syscall.GUID {
	return &IID_HeaderFooter
}

func (this *HeaderFooter) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HeaderFooter) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *HeaderFooter) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HeaderFooter) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *HeaderFooter) Range() *Range {
	retVal, _ := this.PropGet(0x00000000, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *HeaderFooter) Index() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *HeaderFooter) IsHeader() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *HeaderFooter) Exists() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *HeaderFooter) SetExists(rhs bool)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *HeaderFooter) PageNumbers() *PageNumbers {
	retVal, _ := this.PropGet(0x00000005, nil)
	return NewPageNumbers(retVal.IDispatch(), false, true)
}

func (this *HeaderFooter) LinkToPrevious() bool {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *HeaderFooter) SetLinkToPrevious(rhs bool)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *HeaderFooter) Shapes() *Shapes {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewShapes(retVal.IDispatch(), false, true)
}

