package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002096F-0000-0000-C000-000000000046
var IID_FontNames = syscall.GUID{0x0002096F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FontNames struct {
	ole.OleClient
}

func NewFontNames(pDisp *win32.IDispatch, addRef bool, scoped bool) *FontNames {
	p := &FontNames{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FontNamesFromVar(v ole.Variant) *FontNames {
	return NewFontNames(v.PdispValVal(), false, false)
}

func (this *FontNames) IID() *syscall.GUID {
	return &IID_FontNames
}

func (this *FontNames) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FontNames) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *FontNames) ForEach(action func(item string) bool) {
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
		pItem, _ := v.ToString()
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *FontNames) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *FontNames) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *FontNames) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *FontNames) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *FontNames) Item(index int32) string {
	retVal := this.Call(0x00000000, []interface{}{index})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

