package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// FA02A26B-6550-45C5-B6F0-80E757CD3482
var IID_Sources = syscall.GUID{0xFA02A26B, 0x6550, 0x45C5, 
	[8]byte{0xB6, 0xF0, 0x80, 0xE7, 0x57, 0xCD, 0x34, 0x82}}

type Sources struct {
	ole.OleClient
}

func NewSources(pDisp *win32.IDispatch, addRef bool, scoped bool) *Sources {
	 if pDisp == nil {
		return nil;
	}
	p := &Sources{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SourcesFromVar(v ole.Variant) *Sources {
	return NewSources(v.IDispatch(), false, false)
}

func (this *Sources) IID() *syscall.GUID {
	return &IID_Sources
}

func (this *Sources) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Sources) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Sources) ForEach(action func(item *Source) bool) {
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
		pItem := (*Source)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Sources) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Sources) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Sources) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Sources) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Sources) Item(index int32) *Source {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewSource(retVal.IDispatch(), false, true)
}

func (this *Sources) Add(data string)  {
	retVal, _ := this.Call(0x0000006b, []interface{}{data})
	_= retVal
}

