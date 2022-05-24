package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209E5-0000-0000-C000-000000000046
var IID_EmailSignatureEntries = syscall.GUID{0x000209E5, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type EmailSignatureEntries struct {
	ole.OleClient
}

func NewEmailSignatureEntries(pDisp *win32.IDispatch, addRef bool, scoped bool) *EmailSignatureEntries {
	 if pDisp == nil {
		return nil;
	}
	p := &EmailSignatureEntries{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EmailSignatureEntriesFromVar(v ole.Variant) *EmailSignatureEntries {
	return NewEmailSignatureEntries(v.IDispatch(), false, false)
}

func (this *EmailSignatureEntries) IID() *syscall.GUID {
	return &IID_EmailSignatureEntries
}

func (this *EmailSignatureEntries) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *EmailSignatureEntries) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *EmailSignatureEntries) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *EmailSignatureEntries) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *EmailSignatureEntries) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *EmailSignatureEntries) ForEach(action func(item *EmailSignatureEntry) bool) {
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
		pItem := (*EmailSignatureEntry)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *EmailSignatureEntries) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *EmailSignatureEntries) Item(index *ole.Variant) *EmailSignatureEntry {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewEmailSignatureEntry(retVal.IDispatch(), false, true)
}

func (this *EmailSignatureEntries) Add(name string, range_ *Range) *EmailSignatureEntry {
	retVal, _ := this.Call(0x00000065, []interface{}{name, range_})
	return NewEmailSignatureEntry(retVal.IDispatch(), false, true)
}

