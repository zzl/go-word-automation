package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 39709229-56A0-4E29-9112-B31DD067EBFD
var IID_BuildingBlockEntries = syscall.GUID{0x39709229, 0x56A0, 0x4E29, 
	[8]byte{0x91, 0x12, 0xB3, 0x1D, 0xD0, 0x67, 0xEB, 0xFD}}

type BuildingBlockEntries struct {
	ole.OleClient
}

func NewBuildingBlockEntries(pDisp *win32.IDispatch, addRef bool, scoped bool) *BuildingBlockEntries {
	 if pDisp == nil {
		return nil;
	}
	p := &BuildingBlockEntries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BuildingBlockEntriesFromVar(v ole.Variant) *BuildingBlockEntries {
	return NewBuildingBlockEntries(v.IDispatch(), false, false)
}

func (this *BuildingBlockEntries) IID() *syscall.GUID {
	return &IID_BuildingBlockEntries
}

func (this *BuildingBlockEntries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *BuildingBlockEntries) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *BuildingBlockEntries) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *BuildingBlockEntries) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *BuildingBlockEntries) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *BuildingBlockEntries) Item(index *ole.Variant) *BuildingBlock {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewBuildingBlock(retVal.IDispatch(), false, true)
}

var BuildingBlockEntries_Add_OptArgs= []string{
	"Description", "InsertOptions", 
}

func (this *BuildingBlockEntries) Add(name string, type_ int32, category string, range_ *Range, optArgs ...interface{}) *BuildingBlock {
	optArgs = ole.ProcessOptArgs(BuildingBlockEntries_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, []interface{}{name, type_, category, range_}, optArgs...)
	return NewBuildingBlock(retVal.IDispatch(), false, true)
}

