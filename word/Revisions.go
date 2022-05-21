package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020980-0000-0000-C000-000000000046
var IID_Revisions = syscall.GUID{0x00020980, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Revisions struct {
	ole.OleClient
}

func NewRevisions(pDisp *win32.IDispatch, addRef bool, scoped bool) *Revisions {
	p := &Revisions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RevisionsFromVar(v ole.Variant) *Revisions {
	return NewRevisions(v.PdispValVal(), false, false)
}

func (this *Revisions) IID() *syscall.GUID {
	return &IID_Revisions
}

func (this *Revisions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Revisions) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Revisions) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Revisions) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Revisions) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Revisions) ForEach(action func(item *Revision) bool) {
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
		pItem := (*Revision)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Revisions) Count() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Revisions) Item(index int32) *Revision {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewRevision(retVal.PdispValVal(), false, true)
}

func (this *Revisions) AcceptAll()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *Revisions) RejectAll()  {
	retVal := this.Call(0x00000066, nil)
	_= retVal
}

