package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002098A-0000-0000-C000-000000000046
var IID_HeadingStyles = syscall.GUID{0x0002098A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HeadingStyles struct {
	ole.OleClient
}

func NewHeadingStyles(pDisp *win32.IDispatch, addRef bool, scoped bool) *HeadingStyles {
	 if pDisp == nil {
		return nil;
	}
	p := &HeadingStyles{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HeadingStylesFromVar(v ole.Variant) *HeadingStyles {
	return NewHeadingStyles(v.IDispatch(), false, false)
}

func (this *HeadingStyles) IID() *syscall.GUID {
	return &IID_HeadingStyles
}

func (this *HeadingStyles) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HeadingStyles) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *HeadingStyles) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HeadingStyles) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *HeadingStyles) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *HeadingStyles) ForEach(action func(item *HeadingStyle) bool) {
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
		pItem := (*HeadingStyle)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *HeadingStyles) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *HeadingStyles) Item(index int32) *HeadingStyle {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewHeadingStyle(retVal.IDispatch(), false, true)
}

func (this *HeadingStyles) Add(style *ole.Variant, level int16) *HeadingStyle {
	retVal, _ := this.Call(0x00000064, []interface{}{style, level})
	return NewHeadingStyle(retVal.IDispatch(), false, true)
}

