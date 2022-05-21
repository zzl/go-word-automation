package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// EFC71F9C-7F42-4CD4-A7A7-970D7A48CD27
var IID_OMathMatCols = syscall.GUID{0xEFC71F9C, 0x7F42, 0x4CD4, 
	[8]byte{0xA7, 0xA7, 0x97, 0x0D, 0x7A, 0x48, 0xCD, 0x27}}

type OMathMatCols struct {
	ole.OleClient
}

func NewOMathMatCols(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathMatCols {
	p := &OMathMatCols{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathMatColsFromVar(v ole.Variant) *OMathMatCols {
	return NewOMathMatCols(v.PdispValVal(), false, false)
}

func (this *OMathMatCols) IID() *syscall.GUID {
	return &IID_OMathMatCols
}

func (this *OMathMatCols) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathMatCols) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OMathMatCols) ForEach(action func(item *OMathMatCol) bool) {
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
		pItem := (*OMathMatCol)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *OMathMatCols) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathMatCols) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathMatCols) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathMatCols) Count() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *OMathMatCols) Item(index int32) *OMathMatCol {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewOMathMatCol(retVal.PdispValVal(), false, true)
}

var OMathMatCols_Add_OptArgs= []string{
	"BeforeCol", 
}

func (this *OMathMatCols) Add(optArgs ...interface{}) *OMathMatCol {
	optArgs = ole.ProcessOptArgs(OMathMatCols_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000c8, nil, optArgs...)
	return NewOMathMatCol(retVal.PdispValVal(), false, true)
}

