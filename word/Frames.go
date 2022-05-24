package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002092B-0000-0000-C000-000000000046
var IID_Frames = syscall.GUID{0x0002092B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Frames struct {
	ole.OleClient
}

func NewFrames(pDisp *win32.IDispatch, addRef bool, scoped bool) *Frames {
	 if pDisp == nil {
		return nil;
	}
	p := &Frames{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FramesFromVar(v ole.Variant) *Frames {
	return NewFrames(v.IDispatch(), false, false)
}

func (this *Frames) IID() *syscall.GUID {
	return &IID_Frames
}

func (this *Frames) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Frames) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Frames) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Frames) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Frames) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Frames) ForEach(action func(item *Frame) bool) {
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
		pItem := (*Frame)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Frames) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Frames) Item(index int32) *Frame {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewFrame(retVal.IDispatch(), false, true)
}

func (this *Frames) Add(range_ *Range) *Frame {
	retVal, _ := this.Call(0x00000064, []interface{}{range_})
	return NewFrame(retVal.IDispatch(), false, true)
}

func (this *Frames) Delete()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

