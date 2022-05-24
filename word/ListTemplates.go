package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020990-0000-0000-C000-000000000046
var IID_ListTemplates = syscall.GUID{0x00020990, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListTemplates struct {
	ole.OleClient
}

func NewListTemplates(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListTemplates {
	 if pDisp == nil {
		return nil;
	}
	p := &ListTemplates{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListTemplatesFromVar(v ole.Variant) *ListTemplates {
	return NewListTemplates(v.IDispatch(), false, false)
}

func (this *ListTemplates) IID() *syscall.GUID {
	return &IID_ListTemplates
}

func (this *ListTemplates) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListTemplates) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ListTemplates) ForEach(action func(item *ListTemplate) bool) {
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
		pItem := (*ListTemplate)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ListTemplates) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *ListTemplates) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListTemplates) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ListTemplates) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListTemplates) Item(index *ole.Variant) *ListTemplate {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewListTemplate(retVal.IDispatch(), false, true)
}

var ListTemplates_Add_OptArgs= []string{
	"OutlineNumbered", "Name", 
}

func (this *ListTemplates) Add(optArgs ...interface{}) *ListTemplate {
	optArgs = ole.ProcessOptArgs(ListTemplates_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000064, nil, optArgs...)
	return NewListTemplate(retVal.IDispatch(), false, true)
}

