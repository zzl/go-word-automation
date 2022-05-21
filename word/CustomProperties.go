package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// B923FDE1-F08C-11D3-91B0-00105A0A19FD
var IID_CustomProperties = syscall.GUID{0xB923FDE1, 0xF08C, 0x11D3, 
	[8]byte{0x91, 0xB0, 0x00, 0x10, 0x5A, 0x0A, 0x19, 0xFD}}

type CustomProperties struct {
	ole.OleClient
}

func NewCustomProperties(pDisp *win32.IDispatch, addRef bool, scoped bool) *CustomProperties {
	p := &CustomProperties{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CustomPropertiesFromVar(v ole.Variant) *CustomProperties {
	return NewCustomProperties(v.PdispValVal(), false, false)
}

func (this *CustomProperties) IID() *syscall.GUID {
	return &IID_CustomProperties
}

func (this *CustomProperties) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CustomProperties) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CustomProperties) ForEach(action func(item *CustomProperty) bool) {
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
		pItem := (*CustomProperty)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *CustomProperties) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *CustomProperties) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CustomProperties) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CustomProperties) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CustomProperties) Item(index *ole.Variant) *CustomProperty {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewCustomProperty(retVal.PdispValVal(), false, true)
}

func (this *CustomProperties) Add(name string, value string) *CustomProperty {
	retVal := this.Call(0x00000005, []interface{}{name, value})
	return NewCustomProperty(retVal.PdispValVal(), false, true)
}

