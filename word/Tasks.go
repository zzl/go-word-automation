package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020983-0000-0000-C000-000000000046
var IID_Tasks = syscall.GUID{0x00020983, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Tasks struct {
	ole.OleClient
}

func NewTasks(pDisp *win32.IDispatch, addRef bool, scoped bool) *Tasks {
	p := &Tasks{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TasksFromVar(v ole.Variant) *Tasks {
	return NewTasks(v.PdispValVal(), false, false)
}

func (this *Tasks) IID() *syscall.GUID {
	return &IID_Tasks
}

func (this *Tasks) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Tasks) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Tasks) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Tasks) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Tasks) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Tasks) ForEach(action func(item *Task) bool) {
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
		pItem := (*Task)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Tasks) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Tasks) Item(index *ole.Variant) *Task {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewTask(retVal.PdispValVal(), false, true)
}

func (this *Tasks) Exists(name string) bool {
	retVal := this.Call(0x00000002, []interface{}{name})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Tasks) ExitWindows()  {
	retVal := this.Call(0x00000003, nil)
	_= retVal
}

