package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020949-0000-0000-C000-000000000046
var IID_AutoCorrect = syscall.GUID{0x00020949, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type AutoCorrect struct {
	ole.OleClient
}

func NewAutoCorrect(pDisp *win32.IDispatch, addRef bool, scoped bool) *AutoCorrect {
	p := &AutoCorrect{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func AutoCorrectFromVar(v ole.Variant) *AutoCorrect {
	return NewAutoCorrect(v.PdispValVal(), false, false)
}

func (this *AutoCorrect) IID() *syscall.GUID {
	return &IID_AutoCorrect
}

func (this *AutoCorrect) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *AutoCorrect) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrect) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AutoCorrect) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *AutoCorrect) CorrectDays() bool {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectDays(rhs bool)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) CorrectInitialCaps() bool {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectInitialCaps(rhs bool)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) CorrectSentenceCaps() bool {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectSentenceCaps(rhs bool)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) ReplaceText() bool {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetReplaceText(rhs bool)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) Entries() *AutoCorrectEntries {
	retVal := this.PropGet(0x00000006, nil)
	return NewAutoCorrectEntries(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrect) FirstLetterExceptions() *FirstLetterExceptions {
	retVal := this.PropGet(0x00000007, nil)
	return NewFirstLetterExceptions(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrect) FirstLetterAutoAdd() bool {
	retVal := this.PropGet(0x00000008, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetFirstLetterAutoAdd(rhs bool)  {
	retVal := this.PropPut(0x00000008, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) TwoInitialCapsExceptions() *TwoInitialCapsExceptions {
	retVal := this.PropGet(0x00000009, nil)
	return NewTwoInitialCapsExceptions(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrect) TwoInitialCapsAutoAdd() bool {
	retVal := this.PropGet(0x0000000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetTwoInitialCapsAutoAdd(rhs bool)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) CorrectCapsLock() bool {
	retVal := this.PropGet(0x0000000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectCapsLock(rhs bool)  {
	retVal := this.PropPut(0x0000000b, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) CorrectHangulAndAlphabet() bool {
	retVal := this.PropGet(0x0000000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectHangulAndAlphabet(rhs bool)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) HangulAndAlphabetExceptions() *HangulAndAlphabetExceptions {
	retVal := this.PropGet(0x0000000d, nil)
	return NewHangulAndAlphabetExceptions(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrect) HangulAndAlphabetAutoAdd() bool {
	retVal := this.PropGet(0x0000000e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetHangulAndAlphabetAutoAdd(rhs bool)  {
	retVal := this.PropPut(0x0000000e, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) ReplaceTextFromSpellingChecker() bool {
	retVal := this.PropGet(0x0000000f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetReplaceTextFromSpellingChecker(rhs bool)  {
	retVal := this.PropPut(0x0000000f, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) OtherCorrectionsAutoAdd() bool {
	retVal := this.PropGet(0x00000010, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetOtherCorrectionsAutoAdd(rhs bool)  {
	retVal := this.PropPut(0x00000010, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) OtherCorrectionsExceptions() *OtherCorrectionsExceptions {
	retVal := this.PropGet(0x00000011, nil)
	return NewOtherCorrectionsExceptions(retVal.PdispValVal(), false, true)
}

func (this *AutoCorrect) CorrectKeyboardSetting() bool {
	retVal := this.PropGet(0x00000012, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectKeyboardSetting(rhs bool)  {
	retVal := this.PropPut(0x00000012, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) CorrectTableCells() bool {
	retVal := this.PropGet(0x00000013, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectTableCells(rhs bool)  {
	retVal := this.PropPut(0x00000013, []interface{}{rhs})
	_= retVal
}

func (this *AutoCorrect) DisplayAutoCorrectOptions() bool {
	retVal := this.PropGet(0x00000014, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetDisplayAutoCorrectOptions(rhs bool)  {
	retVal := this.PropPut(0x00000014, []interface{}{rhs})
	_= retVal
}

