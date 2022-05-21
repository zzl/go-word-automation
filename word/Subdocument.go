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
	return NewSubdocument(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Subdocument) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Subdocument) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Subdocument) Locked() bool {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Subdocument) SetLocked(rhs bool)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *Subdocument) Range() *Range {
	retVal := this.PropGet(0x00000002, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *Subdocument) Name() string {
	retVal := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Subdocument) Path() string {
	retVal := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Subdocument) HasFile() bool {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Subdocument) Level() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Subdocument) Delete()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

func (this *Subdocument) Split(range_ *Range)  {
	retVal := this.Call(0x00000065, []interface{}{range_})
	_= retVal
}

func (this *Subdocument) Open() *Document {
	retVal := this.Call(0x00000066, nil)
	return NewDocument(retVal.PdispValVal(), false, true)
}

