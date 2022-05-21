package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 09760240-0B89-49F7-A79D-479F24723F56
var IID_XMLNode = syscall.GUID{0x09760240, 0x0B89, 0x49F7, 
	[8]byte{0xA7, 0x9D, 0x47, 0x9F, 0x24, 0x72, 0x3F, 0x56}}

type XMLNode struct {
	ole.OleClient
}

func NewXMLNode(pDisp *win32.IDispatch, addRef bool, scoped bool) *XMLNode {
	p := &XMLNode{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XMLNodeFromVar(v ole.Variant) *XMLNode {
	return NewXMLNode(v.PdispValVal(), false, false)
}

func (this *XMLNode) IID() *syscall.GUID {
	return &IID_XMLNode
}

func (this *XMLNode) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XMLNode) BaseName() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLNode) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XMLNode) Range() *Range {
	retVal := this.PropGet(0x00000001, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) Text() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) SetText(rhs string)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *XMLNode) NamespaceURI() string {
	retVal := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) XML(dataOnly bool) string {
	retVal := this.PropGet(0x00000005, []interface{}{dataOnly})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) NextSibling() *XMLNode {
	retVal := this.PropGet(0x00000006, nil)
	return NewXMLNode(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) PreviousSibling() *XMLNode {
	retVal := this.PropGet(0x00000007, nil)
	return NewXMLNode(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) ParentNode() *XMLNode {
	retVal := this.PropGet(0x00000008, nil)
	return NewXMLNode(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) FirstChild() *XMLNode {
	retVal := this.PropGet(0x00000009, nil)
	return NewXMLNode(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) LastChild() *XMLNode {
	retVal := this.PropGet(0x0000000a, nil)
	return NewXMLNode(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) OwnerDocument() *Document {
	retVal := this.PropGet(0x0000000b, nil)
	return NewDocument(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) NodeType() int32 {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *XMLNode) ChildNodes() *XMLNodes {
	retVal := this.PropGet(0x0000000d, nil)
	return NewXMLNodes(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) Attributes() *XMLNodes {
	retVal := this.PropGet(0x0000000f, nil)
	return NewXMLNodes(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) NodeValue() string {
	retVal := this.PropGet(0x00000010, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) SetNodeValue(rhs string)  {
	retVal := this.PropPut(0x00000010, []interface{}{rhs})
	_= retVal
}

func (this *XMLNode) HasChildNodes() bool {
	retVal := this.PropGet(0x00000011, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XMLNode) SelectSingleNode(xpath string, prefixMapping string, fastSearchSkippingTextNodes bool) *XMLNode {
	retVal := this.Call(0x00000012, []interface{}{xpath, prefixMapping, fastSearchSkippingTextNodes})
	return NewXMLNode(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) SelectNodes(xpath string, prefixMapping string, fastSearchSkippingTextNodes bool) *XMLNodes {
	retVal := this.Call(0x00000013, []interface{}{xpath, prefixMapping, fastSearchSkippingTextNodes})
	return NewXMLNodes(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) ChildNodeSuggestions() *XMLChildNodeSuggestions {
	retVal := this.PropGet(0x00000014, nil)
	return NewXMLChildNodeSuggestions(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) Level() int32 {
	retVal := this.PropGet(0x00000015, nil)
	return retVal.LValVal()
}

func (this *XMLNode) ValidationStatus() int32 {
	retVal := this.PropGet(0x00000016, nil)
	return retVal.LValVal()
}

func (this *XMLNode) SmartTag() *SmartTag {
	retVal := this.PropGet(0x00000017, nil)
	return NewSmartTag(retVal.PdispValVal(), false, true)
}

func (this *XMLNode) ValidationErrorText(advanced bool) string {
	retVal := this.PropGet(0x00000018, []interface{}{advanced})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) PlaceholderText() string {
	retVal := this.PropGet(0x00000019, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) SetPlaceholderText(rhs string)  {
	retVal := this.PropPut(0x00000019, []interface{}{rhs})
	_= retVal
}

func (this *XMLNode) Delete()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

func (this *XMLNode) Copy()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *XMLNode) RemoveChild(childElement *XMLNode)  {
	retVal := this.Call(0x00000066, []interface{}{childElement})
	_= retVal
}

func (this *XMLNode) Cut()  {
	retVal := this.Call(0x00000067, nil)
	_= retVal
}

func (this *XMLNode) Validate()  {
	retVal := this.Call(0x00000068, nil)
	_= retVal
}

func (this *XMLNode) SetValidationError(status int32, errorText *ole.Variant, clearedAutomatically bool)  {
	retVal := this.Call(0x00000069, []interface{}{status, errorText, clearedAutomatically})
	_= retVal
}

func (this *XMLNode) WordOpenXML() string {
	retVal := this.PropGet(0x0000006a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

