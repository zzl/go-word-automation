package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002091A-0000-0000-C000-000000000046
var IID_MailMergeDataFields = syscall.GUID{0x0002091A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type MailMergeDataFields struct {
	ole.OleClient
}

func NewMailMergeDataFields(pDisp *win32.IDispatch, addRef bool, scoped bool) *MailMergeDataFields {
	 if pDisp == nil {
		return nil;
	}
	p := &MailMergeDataFields{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MailMergeDataFieldsFromVar(v ole.Variant) *MailMergeDataFields {
	return NewMailMergeDataFields(v.IDispatch(), false, false)
}

func (this *MailMergeDataFields) IID() *syscall.GUID {
	return &IID_MailMergeDataFields
}

func (this *MailMergeDataFields) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MailMergeDataFields) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *MailMergeDataFields) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataFields) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *MailMergeDataFields) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *MailMergeDataFields) ForEach(action func(item *MailMergeDataField) bool) {
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
		pItem := (*MailMergeDataField)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *MailMergeDataFields) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *MailMergeDataFields) Item(index *ole.Variant) *MailMergeDataField {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewMailMergeDataField(retVal.IDispatch(), false, true)
}

