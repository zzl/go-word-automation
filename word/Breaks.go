package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 16BE9309-D708-4322-BB1A-B056F58D17EA
var IID_Breaks = syscall.GUID{0x16BE9309, 0xD708, 0x4322, 
	[8]byte{0xBB, 0x1A, 0xB0, 0x56, 0xF5, 0x8D, 0x17, 0xEA}}

type Breaks struct {
	ole.OleClient
}

func NewBreaks(pDisp *win32.IDispatch, addRef bool, scoped bool) *Breaks {
	 if pDisp == nil {
		return nil;
	}
	p := &Breaks{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BreaksFromVar(v ole.Variant) *Breaks {
	return NewBreaks(v.IDispatch(), false, false)
}

func (this *Breaks) IID() *syscall.GUID {
	return &IID_Breaks
}

func (this *Breaks) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Breaks) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Breaks) ForEach(action func(item *Break) bool) {
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
		pItem := (*Break)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Breaks) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Breaks) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Breaks) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Breaks) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Breaks) Item(index int32) *Break {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewBreak(retVal.IDispatch(), false, true)
}

