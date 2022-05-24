package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209E9-0000-0000-C000-000000000046
var IID_DiagramNode = syscall.GUID{0x000209E9, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DiagramNode struct {
	ole.OleClient
}

func NewDiagramNode(pDisp *win32.IDispatch, addRef bool, scoped bool) *DiagramNode {
	 if pDisp == nil {
		return nil;
	}
	p := &DiagramNode{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DiagramNodeFromVar(v ole.Variant) *DiagramNode {
	return NewDiagramNode(v.IDispatch(), false, false)
}

func (this *DiagramNode) IID() *syscall.GUID {
	return &IID_DiagramNode
}

func (this *DiagramNode) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DiagramNode) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DiagramNode) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *DiagramNode) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000064, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DiagramNode) Children() *DiagramNodeChildren {
	retVal, _ := this.PropGet(0x00000065, nil)
	return NewDiagramNodeChildren(retVal.IDispatch(), false, true)
}

func (this *DiagramNode) Shape() *Shape {
	retVal, _ := this.PropGet(0x00000066, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *DiagramNode) Root() *DiagramNode {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

func (this *DiagramNode) Diagram() *Diagram {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewDiagram(retVal.IDispatch(), false, true)
}

func (this *DiagramNode) Layout() int32 {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *DiagramNode) SetLayout(rhs int32)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *DiagramNode) TextShape() *Shape {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

var DiagramNode_AddNode_OptArgs= []string{
	"Pos", "NodeType", 
}

func (this *DiagramNode) AddNode(optArgs ...interface{}) *DiagramNode {
	optArgs = ole.ProcessOptArgs(DiagramNode_AddNode_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000a, nil, optArgs...)
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

func (this *DiagramNode) Delete()  {
	retVal, _ := this.Call(0x0000000b, nil)
	_= retVal
}

func (this *DiagramNode) MoveNode(targetNode **DiagramNode, pos int32)  {
	retVal, _ := this.Call(0x0000000c, []interface{}{targetNode, pos})
	_= retVal
}

func (this *DiagramNode) ReplaceNode(targetNode **DiagramNode)  {
	retVal, _ := this.Call(0x0000000d, []interface{}{targetNode})
	_= retVal
}

var DiagramNode_SwapNode_OptArgs= []string{
	"Pos", 
}

func (this *DiagramNode) SwapNode(targetNode **DiagramNode, optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(DiagramNode_SwapNode_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000e, []interface{}{targetNode}, optArgs...)
	_= retVal
}

var DiagramNode_CloneNode_OptArgs= []string{
	"TargetNode", "Pos", 
}

func (this *DiagramNode) CloneNode(copyChildren bool, optArgs ...interface{}) *DiagramNode {
	optArgs = ole.ProcessOptArgs(DiagramNode_CloneNode_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000f, []interface{}{copyChildren}, optArgs...)
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

func (this *DiagramNode) TransferChildren(receivingNode **DiagramNode)  {
	retVal, _ := this.Call(0x00000010, []interface{}{receivingNode})
	_= retVal
}

func (this *DiagramNode) NextNode() *DiagramNode {
	retVal, _ := this.Call(0x00000011, nil)
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

func (this *DiagramNode) PrevNode() *DiagramNode {
	retVal, _ := this.Call(0x00000012, nil)
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

