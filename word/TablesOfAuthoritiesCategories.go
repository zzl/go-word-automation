package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020976-0000-0000-C000-000000000046
var IID_TablesOfAuthoritiesCategories = syscall.GUID{0x00020976, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TablesOfAuthoritiesCategories struct {
	ole.OleClient
}

func NewTablesOfAuthoritiesCategories(pDisp *win32.IDispatch, addRef bool, scoped bool) *TablesOfAuthoritiesCategories {
	 if pDisp == nil {
		return nil;
	}
	p := &TablesOfAuthoritiesCategories{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TablesOfAuthoritiesCategoriesFromVar(v ole.Variant) *TablesOfAuthoritiesCategories {
	return NewTablesOfAuthoritiesCategories(v.IDispatch(), false, false)
}

func (this *TablesOfAuthoritiesCategories) IID() *syscall.GUID {
	return &IID_TablesOfAuthoritiesCategories
}

func (this *TablesOfAuthoritiesCategories) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TablesOfAuthoritiesCategories) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TablesOfAuthoritiesCategories) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TablesOfAuthoritiesCategories) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TablesOfAuthoritiesCategories) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *TablesOfAuthoritiesCategories) ForEach(action func(item *TableOfAuthoritiesCategory) bool) {
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
		pItem := (*TableOfAuthoritiesCategory)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *TablesOfAuthoritiesCategories) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *TablesOfAuthoritiesCategories) Item(index *ole.Variant) *TableOfAuthoritiesCategory {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewTableOfAuthoritiesCategory(retVal.IDispatch(), false, true)
}

