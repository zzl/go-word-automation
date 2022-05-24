package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002097F-0000-0000-C000-000000000046
var IID_AddIns = syscall.GUID{0x0002097F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AddIns struct {
	ole.OleClient
}

func NewAddIns(pDisp *win32.IDispatch, addRef bool, scoped bool) *AddIns {
	 if pDisp == nil {
		return nil;
	}
	p := &AddIns{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AddInsFromVar(v ole.Variant) *AddIns {
	return NewAddIns(v.IDispatch(), false, false)
}

func (this *AddIns) IID() *syscall.GUID {
	return &IID_AddIns
}

func (this *AddIns) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AddIns) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AddIns) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AddIns) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AddIns) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *AddIns) ForEach(action func(item *AddIn) bool) {
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
		pItem := (*AddIn)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *AddIns) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *AddIns) Item(index *ole.Variant) *AddIn {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewAddIn(retVal.IDispatch(), false, true)
}

var AddIns_Add_OptArgs= []string{
	"Install", 
}

func (this *AddIns) Add(fileName string, optArgs ...interface{}) *AddIn {
	optArgs = ole.ProcessOptArgs(AddIns_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000002, []interface{}{fileName}, optArgs...)
	return NewAddIn(retVal.IDispatch(), false, true)
}

func (this *AddIns) Unload(removeFromList bool)  {
	retVal, _ := this.Call(0x00000003, []interface{}{removeFromList})
	_= retVal
}

