package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// E2E8A400-0615-427D-ADCC-CAD39FFEBD42
var IID_Lines = syscall.GUID{0xE2E8A400, 0x0615, 0x427D, 
	[8]byte{0xAD, 0xCC, 0xCA, 0xD3, 0x9F, 0xFE, 0xBD, 0x42}}

type Lines struct {
	ole.OleClient
}

func NewLines(pDisp *win32.IDispatch, addRef bool, scoped bool) *Lines {
	p := &Lines{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LinesFromVar(v ole.Variant) *Lines {
	return NewLines(v.PdispValVal(), false, false)
}

func (this *Lines) IID() *syscall.GUID {
	return &IID_Lines
}

func (this *Lines) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Lines) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Lines) ForEach(action func(item *Line) bool) {
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
		pItem := (*Line)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Lines) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Lines) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Lines) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Lines) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Lines) Item(index int32) *Line {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewLine(retVal.PdispValVal(), false, true)
}

