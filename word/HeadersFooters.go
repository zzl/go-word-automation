package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020984-0000-0000-C000-000000000046
var IID_HeadersFooters = syscall.GUID{0x00020984, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HeadersFooters struct {
	ole.OleClient
}

func NewHeadersFooters(pDisp *win32.IDispatch, addRef bool, scoped bool) *HeadersFooters {
	 if pDisp == nil {
		return nil;
	}
	p := &HeadersFooters{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HeadersFootersFromVar(v ole.Variant) *HeadersFooters {
	return NewHeadersFooters(v.IDispatch(), false, false)
}

func (this *HeadersFooters) IID() *syscall.GUID {
	return &IID_HeadersFooters
}

func (this *HeadersFooters) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HeadersFooters) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *HeadersFooters) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HeadersFooters) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *HeadersFooters) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *HeadersFooters) ForEach(action func(item *HeaderFooter) bool) {
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
		pItem := (*HeaderFooter)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *HeadersFooters) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *HeadersFooters) Item(index int32) *HeaderFooter {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewHeaderFooter(retVal.IDispatch(), false, true)
}

