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
	return NewXMLNamespace(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XMLNamespace) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLNamespace) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XMLNamespace) URI() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNamespace) Location(allUsers bool) string {
	retVal := this.PropGet(0x00000003, []interface{}{allUsers})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNamespace) SetLocation(allUsers bool, rhs string)  {
	retVal := this.PropPut(0x00000003, []interface{}{allUsers, rhs})
	_= retVal
}

func (this *XMLNamespace) Alias(allUsers bool) string {
	retVal := this.PropGet(0x00000004, []interface{}{allUsers})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNamespace) SetAlias(allUsers bool, rhs string)  {
	retVal := this.PropPut(0x00000004, []interface{}{allUsers, rhs})
	_= retVal
}

func (this *XMLNamespace) XSLTransforms() *XSLTransforms {
	retVal := this.PropGet(0x00000005, nil)
	return NewXSLTransforms(retVal.PdispValVal(), false, true)
}

func (this *XMLNamespace) DefaultTransform(allUsers bool) *XSLTransform {
	retVal := this.PropGet(0x00000006, []interface{}{allUsers})
	return NewXSLTransform(retVal.PdispValVal(), false, true)
}

func (this *XMLNamespace) SetDefaultTransform(allUsers bool, rhs *XSLTransform)  {
	retVal := this.PropPut(0x00000006, []interface{}{allUsers, rhs})
	_= retVal
}

func (this *XMLNamespace) AttachToDocument(document *ole.Variant)  {
	retVal := this.Call(0x00000064, []interface{}{document})
	_= retVal
}

func (this *XMLNamespace) Delete()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

