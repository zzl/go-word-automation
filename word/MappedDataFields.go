package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 799A6814-EA41-11D3-87CC-00105AA31A34
var IID_MappedDataFields = syscall.GUID{0x799A6814, 0xEA41, 0x11D3, 
	[8]byte{0x87, 0xCC, 0x00, 0x10, 0x5A, 0xA3, 0x1A, 0x34}}

type MappedDataFields struct {
	ole.OleClient
}

func NewMappedDataFields(pDisp *win32.IDispatch, addRef bool, scoped bool) *MappedDataFields {
	p := &MappedDataFields{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MappedDataFieldsFromVar(v ole.Variant) *MappedDataFields {
	return NewMappedDataFields(v.PdispValVal(), false, false)
}

func (this *MappedDataFields) IID() *syscall.GUID {
	return &IID_MappedDataFields
}

func (this *MappedDataFields) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MappedDataFields) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *MappedDataFields) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MappedDataFields) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *MappedDataFields) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *MappedDataFields) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *MappedDataFields) ForEach(action func(item *MappedDataField) bool) {
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
		pItem := (*MappedDataField)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *MappedDataFields) Item(index int32) *MappedDataField {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewMappedDataField(retVal.PdispValVal(), false, true)
}

