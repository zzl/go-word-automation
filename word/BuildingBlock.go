package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// BFD3FC23-F763-4FF8-826E-1AFBF598A4E7
var IID_BuildingBlock = syscall.GUID{0xBFD3FC23, 0xF763, 0x4FF8, 
	[8]byte{0x82, 0x6E, 0x1A, 0xFB, 0xF5, 0x98, 0xA4, 0xE7}}

type BuildingBlock struct {
	ole.OleClient
}

func NewBuildingBlock(pDisp *win32.IDispatch, addRef bool, scoped bool) *BuildingBlock {
	p := &BuildingBlock{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BuildingBlockFromVar(v ole.Variant) *BuildingBlock {
	return NewBuildingBlock(v.PdispValVal(), false, false)
}

func (this *BuildingBlock) IID() *syscall.GUID {
	return &IID_BuildingBlock
}

func (this *BuildingBlock) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *BuildingBlock) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *BuildingBlock) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *BuildingBlock) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *BuildingBlock) Index() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *BuildingBlock) Name() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *BuildingBlock) SetName(rhs string)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *BuildingBlock) Type() *BuildingBlockType {
	retVal := this.PropGet(0x00000003, nil)
	return NewBuildingBlockType(retVal.PdispValVal(), false, true)
}

func (this *BuildingBlock) Description() string {
	retVal := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *BuildingBlock) SetDescription(rhs string)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *BuildingBlock) ID() string {
	retVal := this.PropGet(0x00000005, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *BuildingBlock) Category() *Category {
	retVal := this.PropGet(0x00000006, nil)
	return NewCategory(retVal.PdispValVal(), false, true)
}

func (this *BuildingBlock) Value() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *BuildingBlock) SetValue(rhs string)  {
	retVal := this.PropPut(0x00000000, []interface{}{rhs})
	_= retVal
}

func (this *BuildingBlock) InsertOptions() int32 {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *BuildingBlock) SetInsertOptions(rhs int32)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *BuildingBlock) Delete()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

var BuildingBlock_Insert_OptArgs= []string{
	"RichText", 
}

func (this *BuildingBlock) Insert(where *Range, optArgs ...interface{}) *Range {
	optArgs = ole.ProcessOptArgs(BuildingBlock_Insert_OptArgs, optArgs)
	retVal := this.Call(0x00000066, []interface{}{where}, optArgs...)
	return NewRange(retVal.PdispValVal(), false, true)
}

