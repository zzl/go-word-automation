package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 356B06EC-4908-42A4-81FC-4B5A51F3483B
var IID_XMLSchemaReferences = syscall.GUID{0x356B06EC, 0x4908, 0x42A4, 
	[8]byte{0x81, 0xFC, 0x4B, 0x5A, 0x51, 0xF3, 0x48, 0x3B}}

type XMLSchemaReferences struct {
	ole.OleClient
}

func NewXMLSchemaReferences(pDisp *win32.IDispatch, addRef bool, scoped bool) *XMLSchemaReferences {
	p := &XMLSchemaReferences{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XMLSchemaReferencesFromVar(v ole.Variant) *XMLSchemaReferences {
	return NewXMLSchemaReferences(v.PdispValVal(), false, false)
}

func (this *XMLSchemaReferences) IID() *syscall.GUID {
	return &IID_XMLSchemaReferences
}

func (this *XMLSchemaReferences) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XMLSchemaReferences) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *XMLSchemaReferences) ForEach(action func(item *XMLSchemaReference) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*XMLSchemaReference)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *XMLSchemaReferences) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *XMLSchemaReferences) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XMLSchemaReferences) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLSchemaReferences) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XMLSchemaReferences) AutomaticValidation() bool {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XMLSchemaReferences) SetAutomaticValidation(rhs bool)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *XMLSchemaReferences) AllowSaveAsXMLWithoutValidation() bool {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XMLSchemaReferences) SetAllowSaveAsXMLWithoutValidation(rhs bool)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *XMLSchemaReferences) HideValidationErrors() bool {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XMLSchemaReferences) SetHideValidationErrors(rhs bool)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *XMLSchemaReferences) IgnoreMixedContent() bool {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XMLSchemaReferences) SetIgnoreMixedContent(rhs bool)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *XMLSchemaReferences) ShowPlaceholderText() bool {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *XMLSchemaReferences) SetShowPlaceholderText(rhs bool)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *XMLSchemaReferences) Item(index *ole.Variant) *XMLSchemaReference {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewXMLSchemaReference(retVal.PdispValVal(), false, true)
}

func (this *XMLSchemaReferences) Validate()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

func (this *XMLSchemaReferences) Add(namespaceURI *ole.Variant, alias *ole.Variant, fileName *ole.Variant, installForAllUsers bool) *XMLSchemaReference {
	retVal := this.Call(0x00000065, []interface{}{namespaceURI, alias, fileName, installForAllUsers})
	return NewXMLSchemaReference(retVal.PdispValVal(), false, true)
}

