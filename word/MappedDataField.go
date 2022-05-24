package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 5D311669-EA51-11D3-87CC-00105AA31A34
var IID_MappedDataField = syscall.GUID{0x5D311669, 0xEA51, 0x11D3, 
	[8]byte{0x87, 0xCC, 0x00, 0x10, 0x5A, 0xA3, 0x1A, 0x34}}

type MappedDataField struct {
	ole.OleClient
}

func NewMappedDataField(pDisp *win32.IDispatch, addRef bool, scoped bool) *MappedDataField {
	 if pDisp == nil {
		return nil;
	}
	p := &MappedDataField{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func MappedDataFieldFromVar(v ole.Variant) *MappedDataField {
	return NewMappedDataField(v.IDispatch(), false, false)
}

func (this *MappedDataField) IID() *syscall.GUID {
	return &IID_MappedDataField
}

func (this *MappedDataField) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *MappedDataField) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *MappedDataField) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *MappedDataField) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *MappedDataField) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *MappedDataField) DataFieldName() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MappedDataField) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MappedDataField) Value() string {
	retVal, _ := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *MappedDataField) DataFieldIndex() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *MappedDataField) SetDataFieldIndex(rhs int32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

