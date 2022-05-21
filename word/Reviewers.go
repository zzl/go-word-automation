package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 12DCDC9A-5418-48A3-BBE6-EB57BAE275E8
var IID_Reviewers = syscall.GUID{0x12DCDC9A, 0x5418, 0x48A3, 
	[8]byte{0xBB, 0xE6, 0xEB, 0x57, 0xBA, 0xE2, 0x75, 0xE8}}

type Reviewers struct {
	ole.OleClient
}

func NewReviewers(pDisp *win32.IDispatch, addRef bool, scoped bool) *Reviewers {
	p := &Reviewers{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ReviewersFromVar(v ole.Variant) *Reviewers {
	return NewReviewers(v.PdispValVal(), false, false)
}

func (this *Reviewers) IID() *syscall.GUID {
	return &IID_Reviewers
}

func (this *Reviewers) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Reviewers) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Reviewers) ForEach(action func(item *Reviewer) bool) {
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
		pItem := (*Reviewer)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Reviewers) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Reviewers) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Reviewers) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Reviewers) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Reviewers) Item(index *ole.Variant) *Reviewer {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewReviewer(retVal.PdispValVal(), false, true)
}

