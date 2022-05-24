package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// D36C1F42-7044-4B9E-9CA3-85919454DB04
var IID_XMLNodes = syscall.GUID{0xD36C1F42, 0x7044, 0x4B9E, 
	[8]byte{0x9C, 0xA3, 0x85, 0x91, 0x94, 0x54, 0xDB, 0x04}}

type XMLNodes struct {
	ole.OleClient
}

func NewXMLNodes(pDisp *win32.IDispatch, addRef bool, scoped bool) *XMLNodes {
	 if pDisp == nil {
		return nil;
	}
	p := &XMLNodes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XMLNodesFromVar(v ole.Variant) *XMLNodes {
	return NewXMLNodes(v.IDispatch(), false, false)
}

func (this *XMLNodes) IID() *syscall.GUID {
	return &IID_XMLNodes
}

func (this *XMLNodes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XMLNodes) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *XMLNodes) ForEach(action func(item *XMLNode) bool) {
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
		pItem := (*XMLNode)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *XMLNodes) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *XMLNodes) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *XMLNodes) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLNodes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *XMLNodes) Item(index int32) *XMLNode {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewXMLNode(retVal.IDispatch(), false, true)
}

var XMLNodes_Add_OptArgs= []string{
	"Range", 
}

func (this *XMLNodes) Add(name string, namespace string, optArgs ...interface{}) *XMLNode {
	optArgs = ole.ProcessOptArgs(XMLNodes_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000064, []interface{}{name, namespace}, optArgs...)
	return NewXMLNode(retVal.IDispatch(), false, true)
}

