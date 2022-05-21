package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 817F99FA-CCC4-4971-8E9D-1238F735AAFF
var IID_BuildingBlockType = syscall.GUID{0x817F99FA, 0xCCC4, 0x4971, 
	[8]byte{0x8E, 0x9D, 0x12, 0x38, 0xF7, 0x35, 0xAA, 0xFF}}

type BuildingBlockType struct {
	ole.OleClient
}

func NewBuildingBlockType(pDisp *win32.IDispatch, addRef bool, scoped bool) *BuildingBlockType {
	p := &BuildingBlockType{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BuildingBlockTypeFromVar(v ole.Variant) *BuildingBlockType {
	return NewBuildingBlockType(v.PdispValVal(), false, false)
}

func (this *BuildingBlockType) IID() *syscall.GUID {
	return &IID_BuildingBlockType
}

func (this *BuildingBlockType) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *BuildingBlockType) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *BuildingBlockType) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *BuildingBlockType) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *BuildingBlockType) Index() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *BuildingBlockType) Name() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *BuildingBlockType) Categories() *Categories {
	retVal := this.PropGet(0x00000014, nil)
	return NewCategories(retVal.PdispValVal(), false, true)
}

