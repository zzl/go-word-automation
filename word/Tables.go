package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002094D-0000-0000-C000-000000000046
var IID_Tables = syscall.GUID{0x0002094D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Tables struct {
	ole.OleClient
}

func NewTables(pDisp *win32.IDispatch, addRef bool, scoped bool) *Tables {
	p := &Tables{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TablesFromVar(v ole.Variant) *Tables {
	return NewTables(v.PdispValVal(), false, false)
}

func (this *Tables) IID() *syscall.GUID {
	return &IID_Tables
}

func (this *Tables) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Tables) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Tables) ForEach(action func(item *Table) bool) {
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
		pItem := (*Table)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Tables) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Tables) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Tables) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Tables) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Tables) Item(index int32) *Table {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewTable(retVal.PdispValVal(), false, true)
}

func (this *Tables) AddOld(range_ *Range, numRows int32, numColumns int32) *Table {
	retVal := this.Call(0x00000004, []interface{}{range_, numRows, numColumns})
	return NewTable(retVal.PdispValVal(), false, true)
}

var Tables_Add_OptArgs= []string{
	"DefaultTableBehavior", "AutoFitBehavior", 
}

func (this *Tables) Add(range_ *Range, numRows int32, numColumns int32, optArgs ...interface{}) *Table {
	optArgs = ole.ProcessOptArgs(Tables_Add_OptArgs, optArgs)
	retVal := this.Call(0x000000c8, []interface{}{range_, numRows, numColumns}, optArgs...)
	return NewTable(retVal.PdispValVal(), false, true)
}

func (this *Tables) NestingLevel() int32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

