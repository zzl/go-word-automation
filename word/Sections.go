package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002095A-0000-0000-C000-000000000046
var IID_Sections = syscall.GUID{0x0002095A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Sections struct {
	ole.OleClient
}

func NewSections(pDisp *win32.IDispatch, addRef bool, scoped bool) *Sections {
	 if pDisp == nil {
		return nil;
	}
	p := &Sections{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SectionsFromVar(v ole.Variant) *Sections {
	return NewSections(v.IDispatch(), false, false)
}

func (this *Sections) IID() *syscall.GUID {
	return &IID_Sections
}

func (this *Sections) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Sections) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Sections) ForEach(action func(item *Section) bool) {
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
		pItem := (*Section)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Sections) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Sections) First() *Section {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewSection(retVal.IDispatch(), false, true)
}

func (this *Sections) Last() *Section {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewSection(retVal.IDispatch(), false, true)
}

func (this *Sections) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Sections) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Sections) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Sections) PageSetup() *PageSetup {
	retVal, _ := this.PropGet(0x0000044d, nil)
	return NewPageSetup(retVal.IDispatch(), false, true)
}

func (this *Sections) SetPageSetup(rhs *PageSetup)  {
	_ = this.PropPut(0x0000044d, []interface{}{rhs})
}

func (this *Sections) Item(index int32) *Section {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewSection(retVal.IDispatch(), false, true)
}

var Sections_Add_OptArgs= []string{
	"Range", "Start", 
}

func (this *Sections) Add(optArgs ...interface{}) *Section {
	optArgs = ole.ProcessOptArgs(Sections_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000005, nil, optArgs...)
	return NewSection(retVal.IDispatch(), false, true)
}

