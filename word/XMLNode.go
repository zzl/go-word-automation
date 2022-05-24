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
	 if pDisp == nil {
		return nil;
	}
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
	return NewXMLNode(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *XMLNode) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLNode) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *XMLNode) Range() *Range {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *XMLNode) Text() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) SetText(rhs string)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *XMLNode) NamespaceURI() string {
	retVal, _ := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var XMLNode_XML_OptArgs= []string{
	"DataOnly", 
}

func (this *XMLNode) XML(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(XMLNode_XML_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000005, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) NextSibling() *XMLNode {
	retVal, _ := this.PropGet(0x00000006, nil)
	return NewXMLNode(retVal.IDispatch(), false, true)
}

func (this *XMLNode) PreviousSibling() *XMLNode {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewXMLNode(retVal.IDispatch(), false, true)
}

func (this *XMLNode) ParentNode() *XMLNode {
	retVal, _ := this.PropGet(0x00000008, nil)
	return NewXMLNode(retVal.IDispatch(), false, true)
}

func (this *XMLNode) FirstChild() *XMLNode {
	retVal, _ := this.PropGet(0x00000009, nil)
	return NewXMLNode(retVal.IDispatch(), false, true)
}

func (this *XMLNode) LastChild() *XMLNode {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewXMLNode(retVal.IDispatch(), false, true)
}

func (this *XMLNode) OwnerDocument() *Document {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return NewDocument(retVal.IDispatch(), false, true)
}

func (this *XMLNode) NodeType() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *XMLNode) ChildNodes() *XMLNodes {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return NewXMLNodes(retVal.IDispatch(), false, true)
}

func (this *XMLNode) Attributes() *XMLNodes {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return NewXMLNodes(retVal.IDispatch(), false, true)
}

func (this *XMLNode) NodeValue() string {
	retVal, _ := this.PropGet(0x00000010, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) SetNodeValue(rhs string)  {
	_ = this.PropPut(0x00000010, []interface{}{rhs})
}

func (this *XMLNode) HasChildNodes() bool {
	retVal, _ := this.PropGet(0x00000011, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var XMLNode_SelectSingleNode_OptArgs= []string{
	"PrefixMapping", "FastSearchSkippingTextNodes", 
}

func (this *XMLNode) SelectSingleNode(xpath string, optArgs ...interface{}) *XMLNode {
	optArgs = ole.ProcessOptArgs(XMLNode_SelectSingleNode_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000012, []interface{}{xpath}, optArgs...)
	return NewXMLNode(retVal.IDispatch(), false, true)
}

var XMLNode_SelectNodes_OptArgs= []string{
	"PrefixMapping", "FastSearchSkippingTextNodes", 
}

func (this *XMLNode) SelectNodes(xpath string, optArgs ...interface{}) *XMLNodes {
	optArgs = ole.ProcessOptArgs(XMLNode_SelectNodes_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000013, []interface{}{xpath}, optArgs...)
	return NewXMLNodes(retVal.IDispatch(), false, true)
}

func (this *XMLNode) ChildNodeSuggestions() *XMLChildNodeSuggestions {
	retVal, _ := this.PropGet(0x00000014, nil)
	return NewXMLChildNodeSuggestions(retVal.IDispatch(), false, true)
}

func (this *XMLNode) Level() int32 {
	retVal, _ := this.PropGet(0x00000015, nil)
	return retVal.LValVal()
}

func (this *XMLNode) ValidationStatus() int32 {
	retVal, _ := this.PropGet(0x00000016, nil)
	return retVal.LValVal()
}

func (this *XMLNode) SmartTag() *SmartTag {
	retVal, _ := this.PropGet(0x00000017, nil)
	return NewSmartTag(retVal.IDispatch(), false, true)
}

var XMLNode_ValidationErrorText_OptArgs= []string{
	"Advanced", 
}

func (this *XMLNode) ValidationErrorText(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(XMLNode_ValidationErrorText_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000018, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) PlaceholderText() string {
	retVal, _ := this.PropGet(0x00000019, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLNode) SetPlaceholderText(rhs string)  {
	_ = this.PropPut(0x00000019, []interface{}{rhs})
}

func (this *XMLNode) Delete()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

func (this *XMLNode) Copy()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *XMLNode) RemoveChild(childElement *XMLNode)  {
	retVal, _ := this.Call(0x00000066, []interface{}{childElement})
	_= retVal
}

func (this *XMLNode) Cut()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

func (this *XMLNode) Validate()  {
	retVal, _ := this.Call(0x00000068, nil)
	_= retVal
}

var XMLNode_SetValidationError_OptArgs= []string{
	"ErrorText", "ClearedAutomatically", 
}

func (this *XMLNode) SetValidationError(status int32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(XMLNode_SetValidationError_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000069, []interface{}{status}, optArgs...)
	_= retVal
}

func (this *XMLNode) WordOpenXML() string {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

