package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209AA-0000-0000-C000-000000000046
var IID_SpellingSuggestions = syscall.GUID{0x000209AA, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SpellingSuggestions struct {
	ole.OleClient
}

func NewSpellingSuggestions(pDisp *win32.IDispatch, addRef bool, scoped bool) *SpellingSuggestions {
	p := &SpellingSuggestions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SpellingSuggestionsFromVar(v ole.Variant) *SpellingSuggestions {
	return NewSpellingSuggestions(v.PdispValVal(), false, false)
}

func (this *SpellingSuggestions) IID() *syscall.GUID {
	return &IID_SpellingSuggestions
}

func (this *SpellingSuggestions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SpellingSuggestions) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SpellingSuggestions) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SpellingSuggestions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SpellingSuggestions) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SpellingSuggestions) ForEach(action func(item *SpellingSuggestion) bool) {
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
		pItem := (*SpellingSuggestion)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *SpellingSuggestions) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *SpellingSuggestions) SpellingErrorType() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *SpellingSuggestions) Item(index int32) *SpellingSuggestion {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewSpellingSuggestion(retVal.PdispValVal(), false, true)
}

