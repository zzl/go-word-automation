package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020930-0000-0000-C000-000000000046
var IID_Fields = syscall.GUID{0x00020930, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Fields struct {
	ole.OleClient
}

func NewFields(pDisp *win32.IDispatch, addRef bool, scoped bool) *Fields {
	 if pDisp == nil {
		return nil;
	}
	p := &Fields{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FieldsFromVar(v ole.Variant) *Fields {
	return NewFields(v.IDispatch(), false, false)
}

func (this *Fields) IID() *syscall.GUID {
	return &IID_Fields
}

func (this *Fields) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Fields) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Fields) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Fields) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Fields) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Fields) Locked() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Fields) SetLocked(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *Fields) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Fields) ForEach(action func(item *Field) bool) {
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
		pItem := (*Field)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Fields) Item(index int32) *Field {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewField(retVal.IDispatch(), false, true)
}

func (this *Fields) ToggleShowCodes()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

func (this *Fields) Update() int32 {
	retVal, _ := this.Call(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Fields) Unlink()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *Fields) UpdateSource()  {
	retVal, _ := this.Call(0x00000068, nil)
	_= retVal
}

var Fields_Add_OptArgs= []string{
	"Type", "Text", "PreserveFormatting", 
}

func (this *Fields) Add(range_ *Range, optArgs ...interface{}) *Field {
	optArgs = ole.ProcessOptArgs(Fields_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000069, []interface{}{range_}, optArgs...)
	return NewField(retVal.IDispatch(), false, true)
}

