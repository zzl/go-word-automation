package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 656BBED7-E82D-4B0A-8F97-EC742BA11FFA
var IID_XMLNamespaces = syscall.GUID{0x656BBED7, 0xE82D, 0x4B0A, 
	[8]byte{0x8F, 0x97, 0xEC, 0x74, 0x2B, 0xA1, 0x1F, 0xFA}}

type XMLNamespaces struct {
	ole.OleClient
}

func NewXMLNamespaces(pDisp *win32.IDispatch, addRef bool, scoped bool) *XMLNamespaces {
	p := &XMLNamespaces{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XMLNamespacesFromVar(v ole.Variant) *XMLNamespaces {
	return NewXMLNamespaces(v.PdispValVal(), false, false)
}

func (this *XMLNamespaces) IID() *syscall.GUID {
	return &IID_XMLNamespaces
}

func (this *XMLNamespaces) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XMLNamespaces) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *XMLNamespaces) ForEach(action func(item *XMLNamespace) bool) {
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
		pItem := (*XMLNamespace)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *XMLNamespaces) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *XMLNamespaces) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XMLNamespaces) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLNamespaces) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XMLNamespaces) Item(index *ole.Variant) *XMLNamespace {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewXMLNamespace(retVal.PdispValVal(), false, true)
}

func (this *XMLNamespaces) Add(path string, namespaceURI *ole.Variant, alias *ole.Variant, installForAllUsers bool) *XMLNamespace {
	retVal := this.Call(0x00000065, []interface{}{path, namespaceURI, alias, installForAllUsers})
	return NewXMLNamespace(retVal.PdispValVal(), false, true)
}

func (this *XMLNamespaces) InstallManifest(path string, installForAllUsers bool)  {
	retVal := this.Call(0x00000066, []interface{}{path, installForAllUsers})
	_= retVal
}

