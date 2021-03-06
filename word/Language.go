package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002096D-0000-0000-C000-000000000046
var IID_Language = syscall.GUID{0x0002096D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Language struct {
	ole.OleClient
}

func NewLanguage(pDisp *win32.IDispatch, addRef bool, scoped bool) *Language {
	 if pDisp == nil {
		return nil;
	}
	p := &Language{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LanguageFromVar(v ole.Variant) *Language {
	return NewLanguage(v.IDispatch(), false, false)
}

func (this *Language) IID() *syscall.GUID {
	return &IID_Language
}

func (this *Language) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Language) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Language) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Language) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Language) ID() int32 {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *Language) NameLocal() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Language) Name() string {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Language) ActiveGrammarDictionary() *Dictionary {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *Language) ActiveHyphenationDictionary() *Dictionary {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *Language) ActiveSpellingDictionary() *Dictionary {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *Language) ActiveThesaurusDictionary() *Dictionary {
	retVal, _ := this.PropGet(0x00000010, nil)
	return NewDictionary(retVal.IDispatch(), false, true)
}

func (this *Language) DefaultWritingStyle() string {
	retVal, _ := this.PropGet(0x00000011, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Language) SetDefaultWritingStyle(rhs string)  {
	_ = this.PropPut(0x00000011, []interface{}{rhs})
}

func (this *Language) WritingStyleList() ole.Variant {
	retVal, _ := this.PropGet(0x00000012, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *Language) SpellingDictionaryType() int32 {
	retVal, _ := this.PropGet(0x00000013, nil)
	return retVal.LValVal()
}

func (this *Language) SetSpellingDictionaryType(rhs int32)  {
	_ = this.PropPut(0x00000013, []interface{}{rhs})
}

