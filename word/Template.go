package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002096A-0000-0000-C000-000000000046
var IID_Template = syscall.GUID{0x0002096A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Template struct {
	ole.OleClient
}

func NewTemplate(pDisp *win32.IDispatch, addRef bool, scoped bool) *Template {
	p := &Template{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TemplateFromVar(v ole.Variant) *Template {
	return NewTemplate(v.PdispValVal(), false, false)
}

func (this *Template) IID() *syscall.GUID {
	return &IID_Template
}

func (this *Template) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Template) Name() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Template) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Template) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Template) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Template) Path() string {
	retVal := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Template) AutoTextEntries() *AutoTextEntries {
	retVal := this.PropGet(0x00000002, nil)
	return NewAutoTextEntries(retVal.PdispValVal(), false, true)
}

func (this *Template) LanguageID() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Template) SetLanguageID(rhs int32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *Template) Saved() bool {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Template) SetSaved(rhs bool)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *Template) Type() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Template) FullName() string {
	retVal := this.PropGet(0x00000007, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Template) BuiltInDocumentProperties() *ole.DispatchClass {
	retVal := this.PropGet(0x00000008, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Template) CustomDocumentProperties() *ole.DispatchClass {
	retVal := this.PropGet(0x00000009, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Template) ListTemplates() *ListTemplates {
	retVal := this.PropGet(0x0000000a, nil)
	return NewListTemplates(retVal.PdispValVal(), false, true)
}

func (this *Template) LanguageIDFarEast() int32 {
	retVal := this.PropGet(0x0000000b, nil)
	return retVal.LValVal()
}

func (this *Template) SetLanguageIDFarEast(rhs int32)  {
	retVal := this.PropPut(0x0000000b, []interface{}{rhs})
	_= retVal
}

func (this *Template) VBProject() *ole.DispatchClass {
	retVal := this.PropGet(0x00000063, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Template) KerningByAlgorithm() bool {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Template) SetKerningByAlgorithm(rhs bool)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *Template) JustificationMode() int32 {
	retVal := this.PropGet(0x0000000d, nil)
	return retVal.LValVal()
}

func (this *Template) SetJustificationMode(rhs int32)  {
	retVal := this.PropPut(0x0000000d, []interface{}{rhs})
	_= retVal
}

func (this *Template) FarEastLineBreakLevel() int32 {
	retVal := this.PropGet(0x0000000e, nil)
	return retVal.LValVal()
}

func (this *Template) SetFarEastLineBreakLevel(rhs int32)  {
	retVal := this.PropPut(0x0000000e, []interface{}{rhs})
	_= retVal
}

func (this *Template) NoLineBreakBefore() string {
	retVal := this.PropGet(0x0000000f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Template) SetNoLineBreakBefore(rhs string)  {
	retVal := this.PropPut(0x0000000f, []interface{}{rhs})
	_= retVal
}

func (this *Template) NoLineBreakAfter() string {
	retVal := this.PropGet(0x00000010, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Template) SetNoLineBreakAfter(rhs string)  {
	retVal := this.PropPut(0x00000010, []interface{}{rhs})
	_= retVal
}

func (this *Template) OpenAsDocument() *Document {
	retVal := this.Call(0x00000064, nil)
	return NewDocument(retVal.PdispValVal(), false, true)
}

func (this *Template) Save()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *Template) NoProofing() int32 {
	retVal := this.PropGet(0x00000011, nil)
	return retVal.LValVal()
}

func (this *Template) SetNoProofing(rhs int32)  {
	retVal := this.PropPut(0x00000011, []interface{}{rhs})
	_= retVal
}

func (this *Template) FarEastLineBreakLanguage() int32 {
	retVal := this.PropGet(0x00000012, nil)
	return retVal.LValVal()
}

func (this *Template) SetFarEastLineBreakLanguage(rhs int32)  {
	retVal := this.PropPut(0x00000012, []interface{}{rhs})
	_= retVal
}

func (this *Template) BuildingBlockEntries() *BuildingBlockEntries {
	retVal := this.PropGet(0x00000013, nil)
	return NewBuildingBlockEntries(retVal.PdispValVal(), false, true)
}

func (this *Template) BuildingBlockTypes() *BuildingBlockTypes {
	retVal := this.PropGet(0x00000014, nil)
	return NewBuildingBlockTypes(retVal.PdispValVal(), false, true)
}

