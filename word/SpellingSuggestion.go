package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209AB-0000-0000-C000-000000000046
var IID_SpellingSuggestion = syscall.GUID{0x000209AB, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SpellingSuggestion struct {
	ole.OleClient
}

func NewSpellingSuggestion(pDisp *win32.IDispatch, addRef bool, scoped bool) *SpellingSuggestion {
	p := &SpellingSuggestion{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SpellingSuggestionFromVar(v ole.Variant) *SpellingSuggestion {
	return NewSpellingSuggestion(v.PdispValVal(), false, false)
}

func (this *SpellingSuggestion) IID() *syscall.GUID {
	return &IID_SpellingSuggestion
}

func (this *SpellingSuggestion) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SpellingSuggestion) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SpellingSuggestion) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SpellingSuggestion) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SpellingSuggestion) Name() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

