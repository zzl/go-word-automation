package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209EB-0000-0000-C000-000000000046
var IID_DiagramNodes = syscall.GUID{0x000209EB, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DiagramNodes struct {
	ole.OleClient
}

func NewDiagramNodes(pDisp *win32.IDispatch, addRef bool, scoped bool) *DiagramNodes {
	 if pDisp == nil {
		return nil;
	}
	p := &DiagramNodes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DiagramNodesFromVar(v ole.Variant) *DiagramNodes {
	return NewDiagramNodes(v.IDispatch(), false, false)
}

func (this *DiagramNodes) IID() *syscall.GUID {
	return &IID_DiagramNodes
}

func (this *DiagramNodes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DiagramNodes) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DiagramNodes) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *DiagramNodes) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *DiagramNodes) ForEach(action func(item *DiagramNode) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*DiagramNode)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *DiagramNodes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000064, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DiagramNodes) Count() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *DiagramNodes) Item(index *ole.Variant) *DiagramNode {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewDiagramNode(retVal.IDispatch(), false, true)
}

func (this *DiagramNodes) SelectAll()  {
	retVal, _ := this.Call(0x0000000a, nil)
	_= retVal
}

