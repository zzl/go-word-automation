package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// FE0971F0-5E60-4985-BCDA-95CB0B8E0308
var IID_XMLSchemaReference = syscall.GUID{0xFE0971F0, 0x5E60, 0x4985, 
	[8]byte{0xBC, 0xDA, 0x95, 0xCB, 0x0B, 0x8E, 0x03, 0x08}}

type XMLSchemaReference struct {
	ole.OleClient
}

func NewXMLSchemaReference(pDisp *win32.IDispatch, addRef bool, scoped bool) *XMLSchemaReference {
	p := &XMLSchemaReference{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XMLSchemaReferenceFromVar(v ole.Variant) *XMLSchemaReference {
	return NewXMLSchemaReference(v.PdispValVal(), false, false)
}

func (this *XMLSchemaReference) IID() *syscall.GUID {
	return &IID_XMLSchemaReference
}

func (this *XMLSchemaReference) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XMLSchemaReference) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XMLSchemaReference) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLSchemaReference) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XMLSchemaReference) NamespaceURI() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLSchemaReference) Location() string {
	retVal := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLSchemaReference) Delete()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

func (this *XMLSchemaReference) Reload()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

