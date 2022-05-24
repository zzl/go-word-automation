package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 873E774B-926A-4CB1-878D-635A45187595
var IID_OMaths = syscall.GUID{0x873E774B, 0x926A, 0x4CB1, 
	[8]byte{0x87, 0x8D, 0x63, 0x5A, 0x45, 0x18, 0x75, 0x95}}

type OMaths struct {
	ole.OleClient
}

func NewOMaths(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMaths {
	 if pDisp == nil {
		return nil;
	}
	p := &OMaths{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathsFromVar(v ole.Variant) *OMaths {
	return NewOMaths(v.IDispatch(), false, false)
}

func (this *OMaths) IID() *syscall.GUID {
	return &IID_OMaths
}

func (this *OMaths) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMaths) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *OMaths) ForEach(action func(item *OMath) bool) {
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
		pItem := (*OMath)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *OMaths) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMaths) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMaths) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMaths) Count() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *OMaths) Item(index int32) *OMath {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMaths) Linearize()  {
	retVal, _ := this.Call(0x000000c8, nil)
	_= retVal
}

func (this *OMaths) BuildUp()  {
	retVal, _ := this.Call(0x000000c9, nil)
	_= retVal
}

func (this *OMaths) Add(range_ *Range) *Range {
	retVal, _ := this.Call(0x000000ca, []interface{}{range_})
	return NewRange(retVal.IDispatch(), false, true)
}

