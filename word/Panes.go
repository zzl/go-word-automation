package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002095F-0000-0000-C000-000000000046
var IID_Panes = syscall.GUID{0x0002095F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Panes struct {
	ole.OleClient
}

func NewPanes(pDisp *win32.IDispatch, addRef bool, scoped bool) *Panes {
	 if pDisp == nil {
		return nil;
	}
	p := &Panes{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PanesFromVar(v ole.Variant) *Panes {
	return NewPanes(v.IDispatch(), false, false)
}

func (this *Panes) IID() *syscall.GUID {
	return &IID_Panes
}

func (this *Panes) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Panes) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Panes) ForEach(action func(item *Pane) bool) {
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
		pItem := (*Pane)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Panes) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Panes) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Panes) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Panes) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Panes) Item(index int32) *Pane {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewPane(retVal.IDispatch(), false, true)
}

var Panes_Add_OptArgs= []string{
	"SplitVertical", 
}

func (this *Panes) Add(optArgs ...interface{}) *Pane {
	optArgs = ole.ProcessOptArgs(Panes_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000003, nil, optArgs...)
	return NewPane(retVal.IDispatch(), false, true)
}

