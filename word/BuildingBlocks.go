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
	 if pDisp == nil {
		return nil;
	}
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
	return NewBuildingBlocks(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *BuildingBlocks) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *BuildingBlocks) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *BuildingBlocks) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *BuildingBlocks) Item(index *ole.Variant) *BuildingBlock {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewBuildingBlock(retVal.IDispatch(), false, true)
}

var BuildingBlocks_Add_OptArgs= []string{
	"Description", "InsertOptions", 
}

func (this *BuildingBlocks) Add(name string, range_ *Range, optArgs ...interface{}) *BuildingBlock {
	optArgs = ole.ProcessOptArgs(BuildingBlocks_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, []interface{}{name, range_}, optArgs...)
	return NewBuildingBlock(retVal.IDispatch(), false, true)
}

