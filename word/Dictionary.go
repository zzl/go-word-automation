package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209AD-0000-0000-C000-000000000046
var IID_Dictionary = syscall.GUID{0x000209AD, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Dictionary struct {
	ole.OleClient
}

func NewDictionary(pDisp *win32.IDispatch, addRef bool, scoped bool) *Dictionary {
	 if pDisp == nil {
		return nil;
	}
	p := &Dictionary{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DictionaryFromVar(v ole.Variant) *Dictionary {
	return NewDictionary(v.IDispatch(), false, false)
}

func (this *Dictionary) IID() *syscall.GUID {
	return &IID_Dictionary
}

func (this *Dictionary) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Dictionary) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Dictionary) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Dictionary) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Dictionary) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Dictionary) Path() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Dictionary) LanguageID() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Dictionary) SetLanguageID(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *Dictionary) ReadOnly() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Dictionary) Type() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Dictionary) LanguageSpecific() bool {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Dictionary) SetLanguageSpecific(rhs bool)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *Dictionary) Delete()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

