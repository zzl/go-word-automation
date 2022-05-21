package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020929-0000-0000-C000-000000000046
var IID_FormFields = syscall.GUID{0x00020929, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FormFields struct {
	ole.OleClient
}

func NewFormFields(pDisp *win32.IDispatch, addRef bool, scoped bool) *FormFields {
	p := &FormFields{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FormFieldsFromVar(v ole.Variant) *FormFields {
	return NewFormFields(v.PdispValVal(), false, false)
}

func (this *FormFields) IID() *syscall.GUID {
	return &IID_FormFields
}

func (this *FormFields) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FormFields) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *FormFields) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *FormFields) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FormFields) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *FormFields) Shaded() bool {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *FormFields) SetShaded(rhs bool)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *FormFields) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *FormFields) ForEach(action func(item *FormField) bool) {
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
		pItem := (*FormField)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *FormFields) Item(index *ole.Variant) *FormField {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewFormField(retVal.PdispValVal(), false, true)
}

func (this *FormFields) Add(range_ *Range, type_ int32) *FormField {
	retVal := this.Call(0x00000065, []interface{}{range_, type_})
	return NewFormField(retVal.PdispValVal(), false, true)
}

