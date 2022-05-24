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
	 if pDisp == nil {
		return nil;
	}
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
	return NewAutoCorrect(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *AutoCorrect) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *AutoCorrect) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *AutoCorrect) CorrectDays() bool {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectDays(rhs bool)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *AutoCorrect) CorrectInitialCaps() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectInitialCaps(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *AutoCorrect) CorrectSentenceCaps() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectSentenceCaps(rhs bool)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *AutoCorrect) ReplaceText() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetReplaceText(rhs bool)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *AutoCorrect) Entries() *AutoCorrectEntries {
	retVal, _ := this.PropGet(0x00000006, nil)
	return NewAutoCorrectEntries(retVal.IDispatch(), false, true)
}

func (this *AutoCorrect) FirstLetterExceptions() *FirstLetterExceptions {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewFirstLetterExceptions(retVal.IDispatch(), false, true)
}

func (this *AutoCorrect) FirstLetterAutoAdd() bool {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetFirstLetterAutoAdd(rhs bool)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *AutoCorrect) TwoInitialCapsExceptions() *TwoInitialCapsExceptions {
	retVal, _ := this.PropGet(0x00000009, nil)
	return NewTwoInitialCapsExceptions(retVal.IDispatch(), false, true)
}

func (this *AutoCorrect) TwoInitialCapsAutoAdd() bool {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetTwoInitialCapsAutoAdd(rhs bool)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *AutoCorrect) CorrectCapsLock() bool {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectCapsLock(rhs bool)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *AutoCorrect) CorrectHangulAndAlphabet() bool {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectHangulAndAlphabet(rhs bool)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *AutoCorrect) HangulAndAlphabetExceptions() *HangulAndAlphabetExceptions {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return NewHangulAndAlphabetExceptions(retVal.IDispatch(), false, true)
}

func (this *AutoCorrect) HangulAndAlphabetAutoAdd() bool {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetHangulAndAlphabetAutoAdd(rhs bool)  {
	_ = this.PropPut(0x0000000e, []interface{}{rhs})
}

func (this *AutoCorrect) ReplaceTextFromSpellingChecker() bool {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetReplaceTextFromSpellingChecker(rhs bool)  {
	_ = this.PropPut(0x0000000f, []interface{}{rhs})
}

func (this *AutoCorrect) OtherCorrectionsAutoAdd() bool {
	retVal, _ := this.PropGet(0x00000010, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetOtherCorrectionsAutoAdd(rhs bool)  {
	_ = this.PropPut(0x00000010, []interface{}{rhs})
}

func (this *AutoCorrect) OtherCorrectionsExceptions() *OtherCorrectionsExceptions {
	retVal, _ := this.PropGet(0x00000011, nil)
	return NewOtherCorrectionsExceptions(retVal.IDispatch(), false, true)
}

func (this *AutoCorrect) CorrectKeyboardSetting() bool {
	retVal, _ := this.PropGet(0x00000012, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectKeyboardSetting(rhs bool)  {
	_ = this.PropPut(0x00000012, []interface{}{rhs})
}

func (this *AutoCorrect) CorrectTableCells() bool {
	retVal, _ := this.PropGet(0x00000013, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetCorrectTableCells(rhs bool)  {
	_ = this.PropPut(0x00000013, []interface{}{rhs})
}

func (this *AutoCorrect) DisplayAutoCorrectOptions() bool {
	retVal, _ := this.PropGet(0x00000014, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *AutoCorrect) SetDisplayAutoCorrectOptions(rhs bool)  {
	_ = this.PropPut(0x00000014, []interface{}{rhs})
}

