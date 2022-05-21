package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 6F9D1F68-06F7-49EF-8902-185E54EB5E87
var IID_OMathAutoCorrect = syscall.GUID{0x6F9D1F68, 0x06F7, 0x49EF, 
	[8]byte{0x89, 0x02, 0x18, 0x5E, 0x54, 0xEB, 0x5E, 0x87}}

type OMathAutoCorrect struct {
	ole.OleClient
}

func NewOMathAutoCorrect(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathAutoCorrect {
	p := &OMathAutoCorrect{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathAutoCorrectFromVar(v ole.Variant) *OMathAutoCorrect {
	return NewOMathAutoCorrect(v.PdispValVal(), false, false)
}

func (this *OMathAutoCorrect) IID() *syscall.GUID {
	return &IID_OMathAutoCorrect
}

func (this *OMathAutoCorrect) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathAutoCorrect) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathAutoCorrect) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathAutoCorrect) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathAutoCorrect) ReplaceText() bool {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathAutoCorrect) SetReplaceText(rhs bool)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *OMathAutoCorrect) UseOutsideOMath() bool {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *OMathAutoCorrect) SetUseOutsideOMath(rhs bool)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *OMathAutoCorrect) Entries() *OMathAutoCorrectEntries {
	retVal := this.PropGet(0x00000069, nil)
	return NewOMathAutoCorrectEntries(retVal.PdispValVal(), false, true)
}

func (this *OMathAutoCorrect) Functions() *OMathRecognizedFunctions {
	retVal := this.PropGet(0x0000006a, nil)
	return NewOMathRecognizedFunctions(retVal.PdispValVal(), false, true)
}

