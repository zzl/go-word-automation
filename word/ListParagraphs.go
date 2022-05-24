package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020991-0000-0000-C000-000000000046
var IID_ListParagraphs = syscall.GUID{0x00020991, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListParagraphs struct {
	ole.OleClient
}

func NewListParagraphs(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListParagraphs {
	 if pDisp == nil {
		return nil;
	}
	p := &ListParagraphs{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListParagraphsFromVar(v ole.Variant) *ListParagraphs {
	return NewListParagraphs(v.IDispatch(), false, false)
}

func (this *ListParagraphs) IID() *syscall.GUID {
	return &IID_ListParagraphs
}

func (this *ListParagraphs) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListParagraphs) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ListParagraphs) ForEach(action func(item *Paragraph) bool) {
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
		pItem := (*Paragraph)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ListParagraphs) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *ListParagraphs) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListParagraphs) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ListParagraphs) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListParagraphs) Item(index int32) *Paragraph {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewParagraph(retVal.IDispatch(), false, true)
}

