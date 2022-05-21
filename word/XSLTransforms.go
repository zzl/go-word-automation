package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// C774F5EA-A539-4284-A1BE-30AEC052D899
var IID_XSLTransforms = syscall.GUID{0xC774F5EA, 0xA539, 0x4284, 
	[8]byte{0xA1, 0xBE, 0x30, 0xAE, 0xC0, 0x52, 0xD8, 0x99}}

type XSLTransforms struct {
	ole.OleClient
}

func NewXSLTransforms(pDisp *win32.IDispatch, addRef bool, scoped bool) *XSLTransforms {
	p := &XSLTransforms{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XSLTransformsFromVar(v ole.Variant) *XSLTransforms {
	return NewXSLTransforms(v.PdispValVal(), false, false)
}

func (this *XSLTransforms) IID() *syscall.GUID {
	return &IID_XSLTransforms
}

func (this *XSLTransforms) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XSLTransforms) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *XSLTransforms) ForEach(action func(item *XSLTransform) bool) {
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
		pItem := (*XSLTransform)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *XSLTransforms) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *XSLTransforms) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XSLTransforms) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XSLTransforms) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XSLTransforms) Item(index *ole.Variant) *XSLTransform {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewXSLTransform(retVal.PdispValVal(), false, true)
}

func (this *XSLTransforms) Add(location string, alias *ole.Variant, installForAllUsers bool) *XSLTransform {
	retVal := this.Call(0x00000065, []interface{}{location, alias, installForAllUsers})
	return NewXSLTransform(retVal.PdispValVal(), false, true)
}

