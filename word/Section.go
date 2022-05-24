package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020959-0000-0000-C000-000000000046
var IID_Section = syscall.GUID{0x00020959, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Section struct {
	ole.OleClient
}

func NewSection(pDisp *win32.IDispatch, addRef bool, scoped bool) *Section {
	 if pDisp == nil {
		return nil;
	}
	p := &Section{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SectionFromVar(v ole.Variant) *Section {
	return NewSection(v.IDispatch(), false, false)
}

func (this *Section) IID() *syscall.GUID {
	return &IID_Section
}

func (this *Section) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Section) Range() *Range {
	retVal, _ := this.PropGet(0x00000000, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Section) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Section) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Section) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Section) PageSetup() *PageSetup {
	retVal, _ := this.PropGet(0x0000044d, nil)
	return NewPageSetup(retVal.IDispatch(), false, true)
}

func (this *Section) SetPageSetup(rhs *PageSetup)  {
	_ = this.PropPut(0x0000044d, []interface{}{rhs})
}

func (this *Section) Headers() *HeadersFooters {
	retVal, _ := this.PropGet(0x00000079, nil)
	return NewHeadersFooters(retVal.IDispatch(), false, true)
}

func (this *Section) Footers() *HeadersFooters {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return NewHeadersFooters(retVal.IDispatch(), false, true)
}

func (this *Section) ProtectedForForms() bool {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Section) SetProtectedForForms(rhs bool)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *Section) Index() int32 {
	retVal, _ := this.PropGet(0x0000007c, nil)
	return retVal.LValVal()
}

func (this *Section) Borders() *Borders {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.IDispatch(), false, true)
}

func (this *Section) SetBorders(rhs *Borders)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

