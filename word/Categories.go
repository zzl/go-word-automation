package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 6E47678B-A879-4E56-8698-3B7CF169FAD4
var IID_Categories = syscall.GUID{0x6E47678B, 0xA879, 0x4E56, 
	[8]byte{0x86, 0x98, 0x3B, 0x7C, 0xF1, 0x69, 0xFA, 0xD4}}

type Categories struct {
	ole.OleClient
}

func NewCategories(pDisp *win32.IDispatch, addRef bool, scoped bool) *Categories {
	p := &Categories{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CategoriesFromVar(v ole.Variant) *Categories {
	return NewCategories(v.PdispValVal(), false, false)
}

func (this *Categories) IID() *syscall.GUID {
	return &IID_Categories
}

func (this *Categories) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Categories) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Categories) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Categories) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Categories) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Categories) Item(index *ole.Variant) *Category {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewCategory(retVal.PdispValVal(), false, true)
}

