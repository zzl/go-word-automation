package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 1B426348-607D-433C-9216-C5D2BF0EF31F
var IID_OMathMatRows = syscall.GUID{0x1B426348, 0x607D, 0x433C, 
	[8]byte{0x92, 0x16, 0xC5, 0xD2, 0xBF, 0x0E, 0xF3, 0x1F}}

type OMathMatRows struct {
	ole.OleClient
}

func NewOMathMatRows(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathMatRows {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathMatRows{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathMatRowsFromVar(v ole.Variant) *OMathMatRows {
	return NewOMathMatRows(v.IDispatch(), false, false)
}

func (this *OMathMatRows) IID() *syscall.GUID {
	return &IID_OMathMatRows
}

func (this *OMathMatRows) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathMatRows) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OMathMatRows) ForEach(action func(item *OMathMatRow) bool) {
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
		pItem := (*OMathMatRow)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *OMathMatRows) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathMatRows) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathMatRows) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathMatRows) Count() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *OMathMatRows) Item(index int32) *OMathMatRow {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewOMathMatRow(retVal.IDispatch(), false, true)
}

var OMathMatRows_Add_OptArgs= []string{
	"BeforeRow", 
}

func (this *OMathMatRows) Add(optArgs ...interface{}) *OMathMatRow {
	optArgs = ole.ProcessOptArgs(OMathMatRows_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000c8, nil, optArgs...)
	return NewOMathMatRow(retVal.IDispatch(), false, true)
}

