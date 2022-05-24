package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020977-0000-0000-C000-000000000046
var IID_TableOfAuthoritiesCategory = syscall.GUID{0x00020977, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TableOfAuthoritiesCategory struct {
	ole.OleClient
}

func NewTableOfAuthoritiesCategory(pDisp *win32.IDispatch, addRef bool, scoped bool) *TableOfAuthoritiesCategory {
	 if pDisp == nil {
		return nil;
	}
	p := &TableOfAuthoritiesCategory{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TableOfAuthoritiesCategoryFromVar(v ole.Variant) *TableOfAuthoritiesCategory {
	return NewTableOfAuthoritiesCategory(v.IDispatch(), false, false)
}

func (this *TableOfAuthoritiesCategory) IID() *syscall.GUID {
	return &IID_TableOfAuthoritiesCategory
}

func (this *TableOfAuthoritiesCategory) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TableOfAuthoritiesCategory) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TableOfAuthoritiesCategory) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TableOfAuthoritiesCategory) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TableOfAuthoritiesCategory) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *TableOfAuthoritiesCategory) SetName(rhs string)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *TableOfAuthoritiesCategory) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

