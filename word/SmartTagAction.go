package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// DFB6AA6C-1068-420F-969D-01280FCC1630
var IID_SmartTagAction = syscall.GUID{0xDFB6AA6C, 0x1068, 0x420F, 
	[8]byte{0x96, 0x9D, 0x01, 0x28, 0x0F, 0xCC, 0x16, 0x30}}

type SmartTagAction struct {
	ole.OleClient
}

func NewSmartTagAction(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTagAction {
	 if pDisp == nil {
		return nil;
	}
	p := &SmartTagAction{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagActionFromVar(v ole.Variant) *SmartTagAction {
	return NewSmartTagAction(v.IDispatch(), false, false)
}

func (this *SmartTagAction) IID() *syscall.GUID {
	return &IID_SmartTagAction
}

func (this *SmartTagAction) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTagAction) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTagAction) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SmartTagAction) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SmartTagAction) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SmartTagAction) Execute()  {
	retVal, _ := this.Call(0x000003eb, nil)
	_= retVal
}

func (this *SmartTagAction) Type() int32 {
	retVal, _ := this.PropGet(0x000003ec, nil)
	return retVal.LValVal()
}

func (this *SmartTagAction) PresentInPane() bool {
	retVal, _ := this.PropGet(0x000003ed, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagAction) ExpandHelp() bool {
	retVal, _ := this.PropGet(0x000003ee, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagAction) SetExpandHelp(rhs bool)  {
	_ = this.PropPut(0x000003ee, []interface{}{rhs})
}

func (this *SmartTagAction) CheckboxState() bool {
	retVal, _ := this.PropGet(0x000003ef, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagAction) SetCheckboxState(rhs bool)  {
	_ = this.PropPut(0x000003ef, []interface{}{rhs})
}

func (this *SmartTagAction) TextboxText() string {
	retVal, _ := this.PropGet(0x000003f0, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTagAction) SetTextboxText(rhs string)  {
	_ = this.PropPut(0x000003f0, []interface{}{rhs})
}

func (this *SmartTagAction) ListSelection() int32 {
	retVal, _ := this.PropGet(0x000003f1, nil)
	return retVal.LValVal()
}

func (this *SmartTagAction) SetListSelection(rhs int32)  {
	_ = this.PropPut(0x000003f1, []interface{}{rhs})
}

func (this *SmartTagAction) RadioGroupSelection() int32 {
	retVal, _ := this.PropGet(0x000003f2, nil)
	return retVal.LValVal()
}

func (this *SmartTagAction) SetRadioGroupSelection(rhs int32)  {
	_ = this.PropPut(0x000003f2, []interface{}{rhs})
}

func (this *SmartTagAction) ExpandDocumentFragment() bool {
	retVal, _ := this.PropGet(0x000003f3, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SmartTagAction) SetExpandDocumentFragment(rhs bool)  {
	_ = this.PropPut(0x000003f3, []interface{}{rhs})
}

func (this *SmartTagAction) ActiveXControl() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003f4, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

