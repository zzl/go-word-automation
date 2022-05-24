package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// B140A023-4850-4DA6-BC5F-CC459C4507BC
var IID_XMLNamespace = syscall.GUID{0xB140A023, 0x4850, 0x4DA6, 
	[8]byte{0xBC, 0x5F, 0xCC, 0x45, 0x9C, 0x45, 0x07, 0xBC}}

type XMLNamespace struct {
	ole.OleClient
}

func NewXMLNamespace(pDisp *win32.IDispatch, addRef bool, scoped bool) *XMLNamespace {
	 if pDisp == nil {
		return nil;
	}
	p := &XMLNamespace{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XMLNamespaceFromVar(v ole.Variant) *XMLNamespace {
	return NewXMLNamespace(v.IDispatch(), false, false)
}

func (this *XMLNamespace) IID() *syscall.GUID {
	return &IID_XMLNamespace
}

func (this *XMLNamespace) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XMLNamespace) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *XMLNamespace) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLNamespace) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *XMLNamespace) URI() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var XMLNamespace_Location_OptArgs= []string{
	"AllUsers", 
}

func (this *XMLNamespace) Location(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(XMLNamespace_Location_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000003, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var XMLNamespace_SetLocation_OptArgs= []string{
	"AllUsers", 
}

func (this *XMLNamespace) SetLocation(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(XMLNamespace_SetLocation_OptArgs, optArgs)
	_ = this.PropPut(0x00000003, nil, optArgs...)
}

var XMLNamespace_Alias_OptArgs= []string{
	"AllUsers", 
}

func (this *XMLNamespace) Alias(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(XMLNamespace_Alias_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000004, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var XMLNamespace_SetAlias_OptArgs= []string{
	"AllUsers", 
}

func (this *XMLNamespace) SetAlias(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(XMLNamespace_SetAlias_OptArgs, optArgs)
	_ = this.PropPut(0x00000004, nil, optArgs...)
}

func (this *XMLNamespace) XSLTransforms() *XSLTransforms {
	retVal, _ := this.PropGet(0x00000005, nil)
	return NewXSLTransforms(retVal.IDispatch(), false, true)
}

var XMLNamespace_DefaultTransform_OptArgs= []string{
	"AllUsers", 
}

func (this *XMLNamespace) DefaultTransform(optArgs ...interface{}) *XSLTransform {
	optArgs = ole.ProcessOptArgs(XMLNamespace_DefaultTransform_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000006, nil, optArgs...)
	return NewXSLTransform(retVal.IDispatch(), false, true)
}

var XMLNamespace_SetDefaultTransform_OptArgs= []string{
	"AllUsers", 
}

func (this *XMLNamespace) SetDefaultTransform(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(XMLNamespace_SetDefaultTransform_OptArgs, optArgs)
	_ = this.PropPut(0x00000006, nil, optArgs...)
}

func (this *XMLNamespace) AttachToDocument(document *ole.Variant)  {
	retVal, _ := this.Call(0x00000064, []interface{}{document})
	_= retVal
}

func (this *XMLNamespace) Delete()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

