package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// B5828B50-0E3D-448A-962D-A40702A5868D
var IID_BuildingBlockTypes = syscall.GUID{0xB5828B50, 0x0E3D, 0x448A, 
	[8]byte{0x96, 0x2D, 0xA4, 0x07, 0x02, 0xA5, 0x86, 0x8D}}

type BuildingBlockTypes struct {
	ole.OleClient
}

func NewBuildingBlockTypes(pDisp *win32.IDispatch, addRef bool, scoped bool) *BuildingBlockTypes {
	 if pDisp == nil {
		return nil;
	}
	p := &BuildingBlockTypes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BuildingBlockTypesFromVar(v ole.Variant) *BuildingBlockTypes {
	return NewBuildingBlockTypes(v.IDispatch(), false, false)
}

func (this *BuildingBlockTypes) IID() *syscall.GUID {
	return &IID_BuildingBlockTypes
}

func (this *BuildingBlockTypes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *BuildingBlockTypes) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *BuildingBlockTypes) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *BuildingBlockTypes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *BuildingBlockTypes) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *BuildingBlockTypes) Item(index int32) *BuildingBlockType {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewBuildingBlockType(retVal.IDispatch(), false, true)
}

