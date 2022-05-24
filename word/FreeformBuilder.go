package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209C9-0000-0000-C000-000000000046
var IID_FreeformBuilder = syscall.GUID{0x000209C9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FreeformBuilder struct {
	ole.OleClient
}

func NewFreeformBuilder(pDisp *win32.IDispatch, addRef bool, scoped bool) *FreeformBuilder {
	 if pDisp == nil {
		return nil;
	}
	p := &FreeformBuilder{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FreeformBuilderFromVar(v ole.Variant) *FreeformBuilder {
	return NewFreeformBuilder(v.IDispatch(), false, false)
}

func (this *FreeformBuilder) IID() *syscall.GUID {
	return &IID_FreeformBuilder
}

func (this *FreeformBuilder) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FreeformBuilder) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *FreeformBuilder) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *FreeformBuilder) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var FreeformBuilder_AddNodes_OptArgs= []string{
	"X2", "Y2", "X3", "Y3", 
}

func (this *FreeformBuilder) AddNodes(segmentType int32, editingType int32, x1 float32, y1 float32, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(FreeformBuilder_AddNodes_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000a, []interface{}{segmentType, editingType, x1, y1}, optArgs...)
	_= retVal
}

var FreeformBuilder_ConvertToShape_OptArgs= []string{
	"Anchor", 
}

func (this *FreeformBuilder) ConvertToShape(optArgs ...interface{}) *Shape {
	optArgs = ole.ProcessOptArgs(FreeformBuilder_ConvertToShape_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000b, nil, optArgs...)
	return NewShape(retVal.IDispatch(), false, true)
}

