package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002097A-0000-0000-C000-000000000046
var IID_AutoCaptions = syscall.GUID{0x0002097A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AutoCaptions struct {
	ole.OleClient
}

func NewAutoCaptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *AutoCaptions {
	 if pDisp == nil {
		return nil;
	}
	p := &AutoCaptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AutoCaptionsFromVar(v ole.Variant) *AutoCaptions {
	return NewAutoCaptions(v.IDispatch(), false, false)
}

func (this *AutoCaptions) IID() *syscall.GUID {
	return &IID_AutoCaptions
}

func (this *AutoCaptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AutoCaptions) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AutoCaptions) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AutoCaptions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AutoCaptions) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *AutoCaptions) ForEach(action func(item *AutoCaption) bool) {
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
		pItem := (*AutoCaption)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *AutoCaptions) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *AutoCaptions) Item(index *ole.Variant) *AutoCaption {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewAutoCaption(retVal.IDispatch(), false, true)
}

func (this *AutoCaptions) CancelAutoInsert()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

