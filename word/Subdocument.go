package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020989-0000-0000-C000-000000000046
var IID_Subdocument = syscall.GUID{0x00020989, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Subdocument struct {
	ole.OleClient
}

func NewSubdocument(pDisp *win32.IDispatch, addRef bool, scoped bool) *Subdocument {
	 if pDisp == nil {
		return nil;
	}
	p := &Subdocument{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SubdocumentFromVar(v ole.Variant) *Subdocument {
	return NewSubdocument(v.IDispatch(), false, false)
}

func (this *Subdocument) IID() *syscall.GUID {
	return &IID_Subdocument
}

func (this *Subdocument) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Subdocument) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Subdocument) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Subdocument) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Subdocument) Locked() bool {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Subdocument) SetLocked(rhs bool)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *Subdocument) Range() *Range {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Subdocument) Name() string {
	retVal, _ := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Subdocument) Path() string {
	retVal, _ := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Subdocument) HasFile() bool {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Subdocument) Level() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Subdocument) Delete()  {
	retVal, _ := this.Call(0x00000064, nil)
	_= retVal
}

func (this *Subdocument) Split(range_ *Range)  {
	retVal, _ := this.Call(0x00000065, []interface{}{range_})
	_= retVal
}

func (this *Subdocument) Open() *Document {
	retVal, _ := this.Call(0x00000066, nil)
	return NewDocument(retVal.IDispatch(), false, true)
}

