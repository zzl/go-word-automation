package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020978-0000-0000-C000-000000000046
var IID_CaptionLabels = syscall.GUID{0x00020978, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CaptionLabels struct {
	ole.OleClient
}

func NewCaptionLabels(pDisp *win32.IDispatch, addRef bool, scoped bool) *CaptionLabels {
	p := &CaptionLabels{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CaptionLabelsFromVar(v ole.Variant) *CaptionLabels {
	return NewCaptionLabels(v.PdispValVal(), false, false)
}

func (this *CaptionLabels) IID() *syscall.GUID {
	return &IID_CaptionLabels
}

func (this *CaptionLabels) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CaptionLabels) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CaptionLabels) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CaptionLabels) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CaptionLabels) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CaptionLabels) ForEach(action func(item *CaptionLabel) bool) {
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
		pItem := (*CaptionLabel)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *CaptionLabels) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *CaptionLabels) Item(index *ole.Variant) *CaptionLabel {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewCaptionLabel(retVal.PdispValVal(), false, true)
}

func (this *CaptionLabels) Add(name string) *CaptionLabel {
	retVal := this.Call(0x00000064, []interface{}{name})
	return NewCaptionLabel(retVal.PdispValVal(), false, true)
}

