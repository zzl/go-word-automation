package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// E6AAEC05-E543-4085-BA92-9BF7D2474F51
var IID_Research = syscall.GUID{0xE6AAEC05, 0xE543, 0x4085, 
	[8]byte{0xBA, 0x92, 0x9B, 0xF7, 0xD2, 0x47, 0x4F, 0x51}}

type Research struct {
	ole.OleClient
}

func NewResearch(pDisp *win32.IDispatch, addRef bool, scoped bool) *Research {
	p := &Research{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ResearchFromVar(v ole.Variant) *Research {
	return NewResearch(v.PdispValVal(), false, false)
}

func (this *Research) IID() *syscall.GUID {
	return &IID_Research
}

func (this *Research) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Research) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Research) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Research) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Research) Query(serviceID string, queryString string, queryLanguage int32, useSelection bool, launchQuery bool) ole.Variant {
	retVal := this.Call(0x000001f4, []interface{}{serviceID, queryString, queryLanguage, useSelection, launchQuery})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Research) SetLanguagePair(languageFrom int32, languageTo int32) ole.Variant {
	retVal := this.Call(0x000001f5, []interface{}{languageFrom, languageTo})
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Research) IsResearchService(serviceID string) bool {
	retVal := this.Call(0x000001f6, []interface{}{serviceID})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Research) FavoriteService() string {
	retVal := this.PropGet(0x000003eb, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Research) SetFavoriteService(rhs string)  {
	retVal := this.PropPut(0x000003eb, []interface{}{rhs})
	_= retVal
}

