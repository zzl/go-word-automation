package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002095C-0000-0000-C000-000000000046
var IID_Words = syscall.GUID{0x0002095C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Words struct {
	ole.OleClient
}

func NewWords(pDisp *win32.IDispatch, addRef bool, scoped bool) *Words {
	 if pDisp == nil {
		return nil;
	}
	p := &Words{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WordsFromVar(v ole.Variant) *Words {
	return NewWords(v.IDispatch(), false, false)
}

func (this *Words) IID() *syscall.GUID {
	return &IID_Words
}

func (this *Words) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Words) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Words) ForEach(action func(item *Range) bool) {
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
		pItem := (*Range)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Words) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Words) First() *Range {
	retVal, _ := this.PropGet(0x00000003, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Words) Last() *Range {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Words) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Words) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Words) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Words) Item(index int32) *Range {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewRange(retVal.IDispatch(), false, true)
}

