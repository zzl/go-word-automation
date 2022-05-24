package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 07B7CC7E-E66C-11D3-9454-00105AA31A08
var IID_StyleSheets = syscall.GUID{0x07B7CC7E, 0xE66C, 0x11D3, 
	[8]byte{0x94, 0x54, 0x00, 0x10, 0x5A, 0xA3, 0x1A, 0x08}}

type StyleSheets struct {
	ole.OleClient
}

func NewStyleSheets(pDisp *win32.IDispatch, addRef bool, scoped bool) *StyleSheets {
	 if pDisp == nil {
		return nil;
	}
	p := &StyleSheets{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func StyleSheetsFromVar(v ole.Variant) *StyleSheets {
	return NewStyleSheets(v.IDispatch(), false, false)
}

func (this *StyleSheets) IID() *syscall.GUID {
	return &IID_StyleSheets
}

func (this *StyleSheets) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *StyleSheets) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *StyleSheets) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *StyleSheets) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *StyleSheets) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *StyleSheets) ForEach(action func(item *StyleSheet) bool) {
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
		pItem := (*StyleSheet)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *StyleSheets) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *StyleSheets) Item(index *ole.Variant) *StyleSheet {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewStyleSheet(retVal.IDispatch(), false, true)
}

func (this *StyleSheets) Add(fileName string, linkType int32, title string, precedence int32) *StyleSheet {
	retVal, _ := this.Call(0x00000002, []interface{}{fileName, linkType, title, precedence})
	return NewStyleSheet(retVal.IDispatch(), false, true)
}

