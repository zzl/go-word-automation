package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// B9F1A4E2-0D0A-43B7-8495-139E7ACBD840
var IID_TaskPane = syscall.GUID{0xB9F1A4E2, 0x0D0A, 0x43B7, 
	[8]byte{0x84, 0x95, 0x13, 0x9E, 0x7A, 0xCB, 0xD8, 0x40}}

type TaskPane struct {
	ole.OleClient
}

func NewTaskPane(pDisp *win32.IDispatch, addRef bool, scoped bool) *TaskPane {
	 if pDisp == nil {
		return nil;
	}
	p := &TaskPane{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TaskPaneFromVar(v ole.Variant) *TaskPane {
	return NewTaskPane(v.IDispatch(), false, false)
}

func (this *TaskPane) IID() *syscall.GUID {
	return &IID_TaskPane
}

func (this *TaskPane) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TaskPane) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TaskPane) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TaskPane) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TaskPane) Visible() bool {
	retVal, _ := this.PropGet(0x000003eb, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TaskPane) SetVisible(rhs bool)  {
	_ = this.PropPut(0x000003eb, []interface{}{rhs})
}

