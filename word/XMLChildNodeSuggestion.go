package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// A87E00E9-3AC3-4B53-ABE3-7379653D0E82
var IID_XMLChildNodeSuggestion = syscall.GUID{0xA87E00E9, 0x3AC3, 0x4B53, 
	[8]byte{0xAB, 0xE3, 0x73, 0x79, 0x65, 0x3D, 0x0E, 0x82}}

type XMLChildNodeSuggestion struct {
	ole.OleClient
}

func NewXMLChildNodeSuggestion(pDisp *win32.IDispatch, addRef bool, scoped bool) *XMLChildNodeSuggestion {
	 if pDisp == nil {
		return nil;
	}
	p := &XMLChildNodeSuggestion{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XMLChildNodeSuggestionFromVar(v ole.Variant) *XMLChildNodeSuggestion {
	return NewXMLChildNodeSuggestion(v.IDispatch(), false, false)
}

func (this *XMLChildNodeSuggestion) IID() *syscall.GUID {
	return &IID_XMLChildNodeSuggestion
}

func (this *XMLChildNodeSuggestion) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XMLChildNodeSuggestion) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *XMLChildNodeSuggestion) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLChildNodeSuggestion) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *XMLChildNodeSuggestion) BaseName() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLChildNodeSuggestion) NamespaceURI() string {
	retVal, _ := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLChildNodeSuggestion) XMLSchemaReference() *XMLSchemaReference {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewXMLSchemaReference(retVal.IDispatch(), false, true)
}

var XMLChildNodeSuggestion_Insert_OptArgs= []string{
	"Range", 
}

func (this *XMLChildNodeSuggestion) Insert(optArgs ...interface{}) *XMLNode {
	optArgs = ole.ProcessOptArgs(XMLChildNodeSuggestion_Insert_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000064, nil, optArgs...)
	return NewXMLNode(retVal.IDispatch(), false, true)
}

