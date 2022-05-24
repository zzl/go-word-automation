package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209E8-0000-0000-C000-000000000046
var IID_HTMLDivisions = syscall.GUID{0x000209E8, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type HTMLDivisions struct {
	ole.OleClient
}

func NewHTMLDivisions(pDisp *win32.IDispatch, addRef bool, scoped bool) *HTMLDivisions {
	 if pDisp == nil {
		return nil;
	}
	p := &HTMLDivisions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HTMLDivisionsFromVar(v ole.Variant) *HTMLDivisions {
	return NewHTMLDivisions(v.IDispatch(), false, false)
}

func (this *HTMLDivisions) IID() *syscall.GUID {
	return &IID_HTMLDivisions
}

func (this *HTMLDivisions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HTMLDivisions) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *HTMLDivisions) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *HTMLDivisions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *HTMLDivisions) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *HTMLDivisions) ForEach(action func(item *HTMLDivision) bool) {
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
		pItem := (*HTMLDivision)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *HTMLDivisions) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *HTMLDivisions) NestingLevel() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

var HTMLDivisions_Add_OptArgs= []string{
	"Range", 
}

func (this *HTMLDivisions) Add(optArgs ...interface{}) *HTMLDivision {
	optArgs = ole.ProcessOptArgs(HTMLDivisions_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, nil, optArgs...)
	return NewHTMLDivision(retVal.IDispatch(), false, true)
}

func (this *HTMLDivisions) Item(index int32) *HTMLDivision {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewHTMLDivision(retVal.IDispatch(), false, true)
}

