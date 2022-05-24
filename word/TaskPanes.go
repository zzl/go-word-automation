package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// E6AAEC05-E543-4085-BA92-9BF7D2474F5C
var IID_TaskPanes = syscall.GUID{0xE6AAEC05, 0xE543, 0x4085, 
	[8]byte{0xBA, 0x92, 0x9B, 0xF7, 0xD2, 0x47, 0x4F, 0x5C}}

type TaskPanes struct {
	ole.OleClient
}

func NewTaskPanes(pDisp *win32.IDispatch, addRef bool, scoped bool) *TaskPanes {
	 if pDisp == nil {
		return nil;
	}
	p := &TaskPanes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TaskPanesFromVar(v ole.Variant) *TaskPanes {
	return NewTaskPanes(v.IDispatch(), false, false)
}

func (this *TaskPanes) IID() *syscall.GUID {
	return &IID_TaskPanes
}

func (this *TaskPanes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TaskPanes) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *TaskPanes) ForEach(action func(item *TaskPane) bool) {
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
		pItem := (*TaskPane)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *TaskPanes) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *TaskPanes) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TaskPanes) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TaskPanes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TaskPanes) Item(index int32) *TaskPane {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewTaskPane(retVal.IDispatch(), false, true)
}

