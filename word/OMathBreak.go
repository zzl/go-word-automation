package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 65E515D5-F50B-4951-8F38-FA6AC8707387
var IID_OMathBreak = syscall.GUID{0x65E515D5, 0xF50B, 0x4951, 
	[8]byte{0x8F, 0x38, 0xFA, 0x6A, 0xC8, 0x70, 0x73, 0x87}}

type OMathBreak struct {
	ole.OleClient
}

func NewOMathBreak(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathBreak {
	p := &OMathBreak{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathBreakFromVar(v ole.Variant) *OMathBreak {
	return NewOMathBreak(v.PdispValVal(), false, false)
}

func (this *OMathBreak) IID() *syscall.GUID {
	return &IID_OMathBreak
}

func (this *OMathBreak) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathBreak) Application() *Application {
	retVal := this.PropGet(0x00000065, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathBreak) Creator() int32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *OMathBreak) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000067, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathBreak) Range() *Range {
	retVal := this.PropGet(0x00000068, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *OMathBreak) AlignAt() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *OMathBreak) SetAlignAt(rhs int32)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *OMathBreak) Delete()  {
	retVal := this.Call(0x0000006a, nil)
	_= retVal
}

