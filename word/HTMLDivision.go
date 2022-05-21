package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209E7-0000-0000-C000-000000000046
var IID_HTMLDivision = syscall.GUID{0x000209E7, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HTMLDivision struct {
	ole.OleClient
}

func NewHTMLDivision(pDisp *win32.IDispatch, addRef bool, scoped bool) *HTMLDivision {
	p := &HTMLDivision{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HTMLDivisionFromVar(v ole.Variant) *HTMLDivision {
	return NewHTMLDivision(v.PdispValVal(), false, false)
}

func (this *HTMLDivision) IID() *syscall.GUID {
	return &IID_HTMLDivision
}

func (this *HTMLDivision) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HTMLDivision) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *HTMLDivision) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HTMLDivision) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *HTMLDivision) Range() *Range {
	retVal := this.PropGet(0x00000001, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *HTMLDivision) Borders() *Borders {
	retVal := this.PropGet(0x0000044c, nil)
	return NewBorders(retVal.PdispValVal(), false, true)
}

func (this *HTMLDivision) LeftIndent() float32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.FltValVal()
}

func (this *HTMLDivision) SetLeftIndent(rhs float32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *HTMLDivision) RightIndent() float32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.FltValVal()
}

func (this *HTMLDivision) SetRightIndent(rhs float32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *HTMLDivision) SpaceBefore() float32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.FltValVal()
}

func (this *HTMLDivision) SetSpaceBefore(rhs float32)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *HTMLDivision) SpaceAfter() float32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.FltValVal()
}

func (this *HTMLDivision) SetSpaceAfter(rhs float32)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *HTMLDivision) HTMLDivisions() *HTMLDivisions {
	retVal := this.PropGet(0x00000007, nil)
	return NewHTMLDivisions(retVal.PdispValVal(), false, true)
}

var HTMLDivision_HTMLDivisionParent_OptArgs= []string{
	"LevelsUp", 
}

func (this *HTMLDivision) HTMLDivisionParent(optArgs ...interface{}) *HTMLDivision {
	optArgs = ole.ProcessOptArgs(HTMLDivision_HTMLDivisionParent_OptArgs, optArgs)
	retVal := this.Call(0x00000008, nil, optArgs...)
	return NewHTMLDivision(retVal.PdispValVal(), false, true)
}

func (this *HTMLDivision) Delete()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

