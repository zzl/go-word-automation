package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020988-0000-0000-C000-000000000046
var IID_Subdocuments = syscall.GUID{0x00020988, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Subdocuments struct {
	ole.OleClient
}

func NewSubdocuments(pDisp *win32.IDispatch, addRef bool, scoped bool) *Subdocuments {
	 if pDisp == nil {
		return nil;
	}
	p := &Subdocuments{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SubdocumentsFromVar(v ole.Variant) *Subdocuments {
	return NewSubdocuments(v.IDispatch(), false, false)
}

func (this *Subdocuments) IID() *syscall.GUID {
	return &IID_Subdocuments
}

func (this *Subdocuments) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Subdocuments) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Subdocuments) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Subdocuments) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Subdocuments) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Subdocuments) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Subdocuments) ForEach(action func(item *Subdocument) bool) {
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
		pItem := (*Subdocument)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Subdocuments) Expanded() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Subdocuments) SetExpanded(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *Subdocuments) Item(index int32) *Subdocument {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewSubdocument(retVal.IDispatch(), false, true)
}

var Subdocuments_AddFromFile_OptArgs= []string{
	"ConfirmConversions", "ReadOnly", "PasswordDocument", "PasswordTemplate", 
	"Revert", "WritePasswordDocument", "WritePasswordTemplate", 
}

func (this *Subdocuments) AddFromFile(name *ole.Variant, optArgs ...interface{}) *Subdocument {
	optArgs = ole.ProcessOptArgs(Subdocuments_AddFromFile_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000064, []interface{}{name}, optArgs...)
	return NewSubdocument(retVal.IDispatch(), false, true)
}

func (this *Subdocuments) AddFromRange(range_ *Range) *Subdocument {
	retVal, _ := this.Call(0x00000065, []interface{}{range_})
	return NewSubdocument(retVal.IDispatch(), false, true)
}

var Subdocuments_Merge_OptArgs= []string{
	"FirstSubdocument", "LastSubdocument", 
}

func (this *Subdocuments) Merge(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Subdocuments_Merge_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, nil, optArgs...)
	_= retVal
}

func (this *Subdocuments) Delete()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

func (this *Subdocuments) Select()  {
	retVal, _ := this.Call(0x00000068, nil)
	_= retVal
}

