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
	 if pDisp == nil {
		return nil;
	}
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
	return NewResearch(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Research) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Research) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var Research_Query_OptArgs= []string{
	"QueryString", "QueryLanguage", "UseSelection", "LaunchQuery", 
}

func (this *Research) Query(serviceID string, optArgs ...interface{}) ole.Variant {
	optArgs = ole.ProcessOptArgs(Research_Query_OptArgs, optArgs)
	retVal, _ := this.Call(0x000001f4, []interface{}{serviceID}, optArgs...)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Research) SetLanguagePair(languageFrom int32, languageTo int32) ole.Variant {
	retVal, _ := this.Call(0x000001f5, []interface{}{languageFrom, languageTo})
	com.AddToScope(retVal)
	return *retVal
}

func (this *Research) IsResearchService(serviceID string) bool {
	retVal, _ := this.Call(0x000001f6, []interface{}{serviceID})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Research) FavoriteService() string {
	retVal, _ := this.PropGet(0x000003eb, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Research) SetFavoriteService(rhs string)  {
	_ = this.PropPut(0x000003eb, []interface{}{rhs})
}

