package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// DE63B5AC-CA4F-46FE-9184-A5719AB9ED5B
var IID_XMLChildNodeSuggestions = syscall.GUID{0xDE63B5AC, 0xCA4F, 0x46FE, 
	[8]byte{0x91, 0x84, 0xA5, 0x71, 0x9A, 0xB9, 0xED, 0x5B}}

type XMLChildNodeSuggestions struct {
	ole.OleClient
}

func NewXMLChildNodeSuggestions(pDisp *win32.IDispatch, addRef bool, scoped bool) *XMLChildNodeSuggestions {
	p := &XMLChildNodeSuggestions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XMLChildNodeSuggestionsFromVar(v ole.Variant) *XMLChildNodeSuggestions {
	return NewXMLChildNodeSuggestions(v.PdispValVal(), false, false)
}

func (this *XMLChildNodeSuggestions) IID() *syscall.GUID {
	return &IID_XMLChildNodeSuggestions
}

func (this *XMLChildNodeSuggestions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XMLChildNodeSuggestions) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *XMLChildNodeSuggestions) ForEach(action func(item *XMLChildNodeSuggestion) bool) {
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
		pItem := (*XMLChildNodeSuggestion)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *XMLChildNodeSuggestions) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *XMLChildNodeSuggestions) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XMLChildNodeSuggestions) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XMLChildNodeSuggestions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XMLChildNodeSuggestions) Item(index *ole.Variant) *XMLChildNodeSuggestion {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewXMLChildNodeSuggestion(retVal.PdispValVal(), false, true)
}

