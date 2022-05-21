package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 91807402-6C6F-47CD-B8FA-C42FEE8EE924
var IID_Pages = syscall.GUID{0x91807402, 0x6C6F, 0x47CD, 
	[8]byte{0xB8, 0xFA, 0xC4, 0x2F, 0xEE, 0x8E, 0xE9, 0x24}}

type Pages struct {
	ole.OleClient
}

func NewPages(pDisp *win32.IDispatch, addRef bool, scoped bool) *Pages {
	p := &Pages{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PagesFromVar(v ole.Variant) *Pages {
	return NewPages(v.PdispValVal(), false, false)
}

func (this *Pages) IID() *syscall.GUID {
	return &IID_Pages
}

func (this *Pages) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Pages) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Pages) ForEach(action func(item *Page) bool) {
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
		pItem := (*Page)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Pages) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Pages) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Pages) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Pages) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Pages) Item(index int32) *Page {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewPage(retVal.PdispValVal(), false, true)
}

