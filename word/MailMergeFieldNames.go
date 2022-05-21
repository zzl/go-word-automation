package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002091C-0000-0000-C000-000000000046
var IID_MailMergeFieldNames = syscall.GUID{0x0002091C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailMergeFieldNames struct {
	ole.OleClient
}

func NewMailMergeFieldNames(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailMergeFieldNames {
	p := &MailMergeFieldNames{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailMergeFieldNamesFromVar(v ole.Variant) *MailMergeFieldNames {
	return NewMailMergeFieldNames(v.PdispValVal(), false, false)
}

func (this *MailMergeFieldNames) IID() *syscall.GUID {
	return &IID_MailMergeFieldNames
}

func (this *MailMergeFieldNames) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailMergeFieldNames) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MailMergeFieldNames) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailMergeFieldNames) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MailMergeFieldNames) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *MailMergeFieldNames) ForEach(action func(item *MailMergeFieldName) bool) {
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
		pItem := (*MailMergeFieldName)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *MailMergeFieldNames) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *MailMergeFieldNames) Item(index *ole.Variant) *MailMergeFieldName {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewMailMergeFieldName(retVal.PdispValVal(), false, true)
}

