package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// C6D50987-25D7-408A-BFF2-90BF86A24E93
var IID_BuildingBlocks = syscall.GUID{0xC6D50987, 0x25D7, 0x408A, 
	[8]byte{0xBF, 0xF2, 0x90, 0xBF, 0x86, 0xA2, 0x4E, 0x93}}

type BuildingBlocks struct {
	ole.OleClient
}

func NewBuildingBlocks(pDisp *win32.IDispatch, addRef bool, scoped bool) *BuildingBlocks {
	p := &BuildingBlocks{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BuildingBlocksFromVar(v ole.Variant) *BuildingBlocks {
	return NewBuildingBlocks(v.PdispValVal(), false, false)
}

func (this *BuildingBlocks) IID() *syscall.GUID {
	return &IID_BuildingBlocks
}

func (this *BuildingBlocks) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *BuildingBlocks) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *BuildingBlocks) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *BuildingBlocks) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *BuildingBlocks) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *BuildingBlocks) Item(index *ole.Variant) *BuildingBlock {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewBuildingBlock(retVal.PdispValVal(), false, true)
}

func (this *BuildingBlocks) Add(name string, range_ *Range, description *ole.Variant, insertOptions int32) *BuildingBlock {
	retVal := this.Call(0x00000065, []interface{}{name, range_, description, insertOptions})
	return NewBuildingBlock(retVal.PdispValVal(), false, true)
}

