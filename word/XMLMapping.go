package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0C1FABE7-F737-406F-9CA3-B07661F9D1A2
var IID_XMLMapping = syscall.GUID{0x0C1FABE7, 0xF737, 0x406F, 
	[8]byte{0x9C, 0xA3, 0xB0, 0x76, 0x61, 0xF9, 0xD1, 0xA2}}

type XMLMapping struct {
	ole.OleClient
}

func NewXMLMapping(pDisp *win32.IDispatch, addRef bool, scoped bool) *XMLMapping {
	 if pDisp == nil {
		return nil;
	}
	p := &XMLMapping{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XMLMappingFromVar(v ole.Variant) *XMLMapping {
	return NewXMLMapping(v.IDispatch(), false, false)
}

func (this *XMLMapping) IID() *syscall.GUID {
	return &IID_XMLMapping
}

func (this *XMLMapping) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XMLMapping) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *XMLMapping) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLMapping) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *XMLMapping) IsMapped() bool {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XMLMapping) CustomXMLPart() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *XMLMapping) CustomXMLNode() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000002, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var XMLMapping_SetMapping_OptArgs= []string{
	"PrefixMapping", "Source", 
}

func (this *XMLMapping) SetMapping(xpath string, optArgs ...interface{}) bool {
	optArgs = ole.ProcessOptArgs(XMLMapping_SetMapping_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000003, []interface{}{xpath}, optArgs...)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XMLMapping) Delete()  {
	retVal, _ := this.Call(0x00000004, nil)
	_= retVal
}

func (this *XMLMapping) SetMappingByNode(node *win32.IDispatch) bool {
	retVal, _ := this.Call(0x00000005, []interface{}{node})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XMLMapping) XPath() string {
	retVal, _ := this.PropGet(0x00000006, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XMLMapping) PrefixMappings() string {
	retVal, _ := this.PropGet(0x00000007, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

