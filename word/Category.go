package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// ECFBDB5E-ACD2-4530-AD79-4560B7FF055C
var IID_Category = syscall.GUID{0xECFBDB5E, 0xACD2, 0x4530, 
	[8]byte{0xAD, 0x79, 0x45, 0x60, 0xB7, 0xFF, 0x05, 0x5C}}

type Category struct {
	ole.OleClient
}

func NewCategory(pDisp *win32.IDispatch, addRef bool, scoped bool) *Category {
	 if pDisp == nil {
		return nil;
	}
	p := &Category{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CategoryFromVar(v ole.Variant) *Category {
	return NewCategory(v.IDispatch(), false, false)
}

func (this *Category) IID() *syscall.GUID {
	return &IID_Category
}

func (this *Category) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Category) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Category) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Category) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Category) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Category) Name() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Category) BuildingBlocks() *BuildingBlocks {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewBuildingBlocks(retVal.IDispatch(), false, true)
}

func (this *Category) Type() *BuildingBlockType {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewBuildingBlockType(retVal.IDispatch(), false, true)
}

