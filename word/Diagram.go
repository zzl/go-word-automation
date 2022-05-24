package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209EC-0000-0000-C000-000000000046
var IID_Diagram = syscall.GUID{0x000209EC, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Diagram struct {
	ole.OleClient
}

func NewDiagram(pDisp *win32.IDispatch, addRef bool, scoped bool) *Diagram {
	 if pDisp == nil {
		return nil;
	}
	p := &Diagram{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DiagramFromVar(v ole.Variant) *Diagram {
	return NewDiagram(v.IDispatch(), false, false)
}

func (this *Diagram) IID() *syscall.GUID {
	return &IID_Diagram
}

func (this *Diagram) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Diagram) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Diagram) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Diagram) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000064, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Diagram) Nodes() *DiagramNodes {
	retVal, _ := this.PropGet(0x00000065, nil)
	return NewDiagramNodes(retVal.IDispatch(), false, true)
}

func (this *Diagram) Type() int32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *Diagram) AutoLayout() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *Diagram) SetAutoLayout(rhs int32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *Diagram) Reverse() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *Diagram) SetReverse(rhs int32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *Diagram) AutoFormat() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *Diagram) SetAutoFormat(rhs int32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *Diagram) Convert(type_ int32)  {
	retVal, _ := this.Call(0x0000000a, []interface{}{type_})
	_= retVal
}

func (this *Diagram) FitText()  {
	retVal, _ := this.Call(0x0000000b, nil)
	_= retVal
}

