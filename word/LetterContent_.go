package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209A1-0000-0000-C000-000000000046
var IID_LetterContent_ = syscall.GUID{0x000209A1, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type LetterContent_ struct {
	ole.OleClient
}

func NewLetterContent_(pDisp *win32.IDispatch, addRef bool, scoped bool) *LetterContent_ {
	 if pDisp == nil {
		return nil;
	}
	p := &LetterContent_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LetterContent_FromVar(v ole.Variant) *LetterContent_ {
	return NewLetterContent_(v.IDispatch(), false, false)
}

func (this *LetterContent_) IID() *syscall.GUID {
	return &IID_LetterContent_
}

func (this *LetterContent_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *LetterContent_) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *LetterContent_) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *LetterContent_) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *LetterContent_) Duplicate() *LetterContent {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return NewLetterContent(retVal.IDispatch(), false, true)
}

func (this *LetterContent_) DateFormat() string {
	retVal, _ := this.PropGet(0x00000065, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetDateFormat(rhs string)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *LetterContent_) IncludeHeaderFooter() bool {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *LetterContent_) SetIncludeHeaderFooter(rhs bool)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *LetterContent_) PageDesign() string {
	retVal, _ := this.PropGet(0x00000067, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetPageDesign(rhs string)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *LetterContent_) LetterStyle() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *LetterContent_) SetLetterStyle(rhs int32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

func (this *LetterContent_) Letterhead() bool {
	retVal, _ := this.PropGet(0x00000069, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *LetterContent_) SetLetterhead(rhs bool)  {
	_ = this.PropPut(0x00000069, []interface{}{rhs})
}

func (this *LetterContent_) LetterheadLocation() int32 {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *LetterContent_) SetLetterheadLocation(rhs int32)  {
	_ = this.PropPut(0x0000006a, []interface{}{rhs})
}

func (this *LetterContent_) LetterheadSize() float32 {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.FltValVal()
}

func (this *LetterContent_) SetLetterheadSize(rhs float32)  {
	_ = this.PropPut(0x0000006b, []interface{}{rhs})
}

func (this *LetterContent_) RecipientName() string {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetRecipientName(rhs string)  {
	_ = this.PropPut(0x0000006c, []interface{}{rhs})
}

func (this *LetterContent_) RecipientAddress() string {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetRecipientAddress(rhs string)  {
	_ = this.PropPut(0x0000006d, []interface{}{rhs})
}

func (this *LetterContent_) Salutation() string {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetSalutation(rhs string)  {
	_ = this.PropPut(0x0000006e, []interface{}{rhs})
}

func (this *LetterContent_) SalutationType() int32 {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return retVal.LValVal()
}

func (this *LetterContent_) SetSalutationType(rhs int32)  {
	_ = this.PropPut(0x0000006f, []interface{}{rhs})
}

func (this *LetterContent_) RecipientReference() string {
	retVal, _ := this.PropGet(0x00000070, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetRecipientReference(rhs string)  {
	_ = this.PropPut(0x00000070, []interface{}{rhs})
}

func (this *LetterContent_) MailingInstructions() string {
	retVal, _ := this.PropGet(0x00000072, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetMailingInstructions(rhs string)  {
	_ = this.PropPut(0x00000072, []interface{}{rhs})
}

func (this *LetterContent_) AttentionLine() string {
	retVal, _ := this.PropGet(0x00000073, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetAttentionLine(rhs string)  {
	_ = this.PropPut(0x00000073, []interface{}{rhs})
}

func (this *LetterContent_) Subject() string {
	retVal, _ := this.PropGet(0x00000074, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetSubject(rhs string)  {
	_ = this.PropPut(0x00000074, []interface{}{rhs})
}

func (this *LetterContent_) EnclosureNumber() int32 {
	retVal, _ := this.PropGet(0x00000075, nil)
	return retVal.LValVal()
}

func (this *LetterContent_) SetEnclosureNumber(rhs int32)  {
	_ = this.PropPut(0x00000075, []interface{}{rhs})
}

func (this *LetterContent_) CCList() string {
	retVal, _ := this.PropGet(0x00000076, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetCCList(rhs string)  {
	_ = this.PropPut(0x00000076, []interface{}{rhs})
}

func (this *LetterContent_) ReturnAddress() string {
	retVal, _ := this.PropGet(0x00000077, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetReturnAddress(rhs string)  {
	_ = this.PropPut(0x00000077, []interface{}{rhs})
}

func (this *LetterContent_) SenderName() string {
	retVal, _ := this.PropGet(0x00000078, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetSenderName(rhs string)  {
	_ = this.PropPut(0x00000078, []interface{}{rhs})
}

func (this *LetterContent_) Closing() string {
	retVal, _ := this.PropGet(0x00000079, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetClosing(rhs string)  {
	_ = this.PropPut(0x00000079, []interface{}{rhs})
}

func (this *LetterContent_) SenderCompany() string {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetSenderCompany(rhs string)  {
	_ = this.PropPut(0x0000007a, []interface{}{rhs})
}

func (this *LetterContent_) SenderJobTitle() string {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetSenderJobTitle(rhs string)  {
	_ = this.PropPut(0x0000007b, []interface{}{rhs})
}

func (this *LetterContent_) SenderInitials() string {
	retVal, _ := this.PropGet(0x0000007c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetSenderInitials(rhs string)  {
	_ = this.PropPut(0x0000007c, []interface{}{rhs})
}

func (this *LetterContent_) InfoBlock() bool {
	retVal, _ := this.PropGet(0x0000007d, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *LetterContent_) SetInfoBlock(rhs bool)  {
	_ = this.PropPut(0x0000007d, []interface{}{rhs})
}

func (this *LetterContent_) RecipientCode() string {
	retVal, _ := this.PropGet(0x0000007e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetRecipientCode(rhs string)  {
	_ = this.PropPut(0x0000007e, []interface{}{rhs})
}

func (this *LetterContent_) RecipientGender() int32 {
	retVal, _ := this.PropGet(0x0000007f, nil)
	return retVal.LValVal()
}

func (this *LetterContent_) SetRecipientGender(rhs int32)  {
	_ = this.PropPut(0x0000007f, []interface{}{rhs})
}

func (this *LetterContent_) ReturnAddressShortForm() string {
	retVal, _ := this.PropGet(0x00000080, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetReturnAddressShortForm(rhs string)  {
	_ = this.PropPut(0x00000080, []interface{}{rhs})
}

func (this *LetterContent_) SenderCity() string {
	retVal, _ := this.PropGet(0x00000081, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetSenderCity(rhs string)  {
	_ = this.PropPut(0x00000081, []interface{}{rhs})
}

func (this *LetterContent_) SenderCode() string {
	retVal, _ := this.PropGet(0x00000082, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetSenderCode(rhs string)  {
	_ = this.PropPut(0x00000082, []interface{}{rhs})
}

func (this *LetterContent_) SenderGender() int32 {
	retVal, _ := this.PropGet(0x00000083, nil)
	return retVal.LValVal()
}

func (this *LetterContent_) SetSenderGender(rhs int32)  {
	_ = this.PropPut(0x00000083, []interface{}{rhs})
}

func (this *LetterContent_) SenderReference() string {
	retVal, _ := this.PropGet(0x00000084, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *LetterContent_) SetSenderReference(rhs string)  {
	_ = this.PropPut(0x00000084, []interface{}{rhs})
}

