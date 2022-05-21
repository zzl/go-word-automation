package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 256B6ABA-6A38-4D39-971C-91FDA9922814
var IID_CoAuthors = syscall.GUID{0x256B6ABA, 0x6A38, 0x4D39, 
	[8]byte{0x97, 0x1C, 0x91, 0xFD, 0xA9, 0x92, 0x28, 0x14}}

type CoAuthors struct {
	ole.OleClient
}

func NewCoAuthors(pDisp *win32.IDispatch, addRef bool, scoped bool) *CoAuthors {
	p := &CoAuthors{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CoAuthorsFromVar(v ole.Variant) *CoAuthors {
	return NewCoAuthors(v.PdispValVal(), false, false)
}

func (this *CoAuthors) IID() *syscall.GUID {
	return &IID_CoAuthors
}

func (this *CoAuthors) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CoAuthors) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CoAuthors) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CoAuthors) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CoAuthors) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *CoAuthors) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *CoAuthors) ForEach(action func(item *CoAuthor) bool) {
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
		pItem := (*CoAuthor)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *CoAuthors) Item(index interface{}) *CoAuthor {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewCoAuthor(retVal.PdispValVal(), false, true)
}

