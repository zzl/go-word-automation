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
	return NewSection(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x00000000, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Section) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Section) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Section) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Section) PageSetup() *PageSetup {
	retVal := this.PropGet(0x0000044d, nil)
	return NewPageSetup(retVal.PdispValVal(), false, true)
}

func (this *Section) SetPageSetup(rhs *PageSetup)  {
	retVal := this.PropPut(0x0000044d, []interface{}{rhs})
	_= retVal
}

func (this *Section) Headers() *HeadersFooters {
	retVal := this.PropGet(0x00000079, nil)
	return NewHeadersFooters(retVal.PdispValVal(), false, true)
}

func (this *Section) Footers() *HeadersFooters {
	retVal := this.PropGet(0x0000007a, nil)
	return NewHeadersFooters(retVal.PdispValVal(), false, true)
}

func (this *Section) ProtectedForForms() bool {
	retVal := this.PropGet(0x0000007b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Section) SetProtectedForForms(rhs bool)  {
	retVal := this.PropPut(0x0000007b, []interface{}{rhs})
	_= retVal
}

func (this *Section) Index() int32 {
	retVal := this.PropGet(0x0000007c, nil)
	return retVal.LValVal()
}

func (this *Section) Borders() *Borders {
	retVal := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *Section) SetBorders(rhs *Borders)  {
	retVal := this.PropPut(0x0000044c, []interface{}{rhs})
	_= retVal
}

