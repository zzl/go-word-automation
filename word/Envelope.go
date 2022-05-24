package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020918-0000-0000-C000-000000000046
var IID_Envelope = syscall.GUID{0x00020918, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Envelope struct {
	ole.OleClient
}

func NewEnvelope(pDisp *win32.IDispatch, addRef bool, scoped bool) *Envelope {
	 if pDisp == nil {
		return nil;
	}
	p := &Envelope{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EnvelopeFromVar(v ole.Variant) *Envelope {
	return NewEnvelope(v.IDispatch(), false, false)
}

func (this *Envelope) IID() *syscall.GUID {
	return &IID_Envelope
}

func (this *Envelope) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Envelope) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Envelope) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Envelope) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Envelope) Address() *Range {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Envelope) ReturnAddress() *Range {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Envelope) DefaultPrintBarCode() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Envelope) SetDefaultPrintBarCode(rhs bool)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *Envelope) DefaultPrintFIMA() bool {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Envelope) SetDefaultPrintFIMA(rhs bool)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *Envelope) DefaultHeight() float32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetDefaultHeight(rhs float32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *Envelope) DefaultWidth() float32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetDefaultWidth(rhs float32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *Envelope) DefaultSize() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Envelope) SetDefaultSize(rhs string)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *Envelope) DefaultOmitReturnAddress() bool {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Envelope) SetDefaultOmitReturnAddress(rhs bool)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *Envelope) FeedSource() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *Envelope) SetFeedSource(rhs int32)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *Envelope) AddressFromLeft() float32 {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetAddressFromLeft(rhs float32)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *Envelope) AddressFromTop() float32 {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetAddressFromTop(rhs float32)  {
	_ = this.PropPut(0x0000000e, []interface{}{rhs})
}

func (this *Envelope) ReturnAddressFromLeft() float32 {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetReturnAddressFromLeft(rhs float32)  {
	_ = this.PropPut(0x0000000f, []interface{}{rhs})
}

func (this *Envelope) ReturnAddressFromTop() float32 {
	retVal, _ := this.PropGet(0x00000010, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetReturnAddressFromTop(rhs float32)  {
	_ = this.PropPut(0x00000010, []interface{}{rhs})
}

func (this *Envelope) AddressStyle() *Style {
	retVal, _ := this.PropGet(0x00000011, nil)
	return NewStyle(retVal.IDispatch(), false, true)
}

func (this *Envelope) ReturnAddressStyle() *Style {
	retVal, _ := this.PropGet(0x00000012, nil)
	return NewStyle(retVal.IDispatch(), false, true)
}

func (this *Envelope) DefaultOrientation() int32 {
	retVal, _ := this.PropGet(0x00000013, nil)
	return retVal.LValVal()
}

func (this *Envelope) SetDefaultOrientation(rhs int32)  {
	_ = this.PropPut(0x00000013, []interface{}{rhs})
}

func (this *Envelope) DefaultFaceUp() bool {
	retVal, _ := this.PropGet(0x00000014, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Envelope) SetDefaultFaceUp(rhs bool)  {
	_ = this.PropPut(0x00000014, []interface{}{rhs})
}

var Envelope_Insert2000_OptArgs= []string{
	"ExtractAddress", "Address", "AutoText", "OmitReturnAddress", 
	"ReturnAddress", "ReturnAutoText", "PrintBarCode", "PrintFIMA", 
	"Size", "Height", "Width", "FeedSource", 
	"AddressFromLeft", "AddressFromTop", "ReturnAddressFromLeft", "ReturnAddressFromTop", 
	"DefaultFaceUp", "DefaultOrientation", 
}

func (this *Envelope) Insert2000(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Envelope_Insert2000_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, nil, optArgs...)
	_= retVal
}

var Envelope_PrintOut2000_OptArgs= []string{
	"ExtractAddress", "Address", "AutoText", "OmitReturnAddress", 
	"ReturnAddress", "ReturnAutoText", "PrintBarCode", "PrintFIMA", 
	"Size", "Height", "Width", "FeedSource", 
	"AddressFromLeft", "AddressFromTop", "ReturnAddressFromLeft", "ReturnAddressFromTop", 
	"DefaultFaceUp", "DefaultOrientation", 
}

func (this *Envelope) PrintOut2000(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Envelope_PrintOut2000_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000066, nil, optArgs...)
	_= retVal
}

func (this *Envelope) UpdateDocument()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

func (this *Envelope) Options()  {
	retVal, _ := this.Call(0x00000068, nil)
	_= retVal
}

func (this *Envelope) Vertical() bool {
	retVal, _ := this.PropGet(0x00000016, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Envelope) SetVertical(rhs bool)  {
	_ = this.PropPut(0x00000016, []interface{}{rhs})
}

func (this *Envelope) RecipientNamefromLeft() float32 {
	retVal, _ := this.PropGet(0x00000017, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetRecipientNamefromLeft(rhs float32)  {
	_ = this.PropPut(0x00000017, []interface{}{rhs})
}

func (this *Envelope) RecipientNamefromTop() float32 {
	retVal, _ := this.PropGet(0x00000018, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetRecipientNamefromTop(rhs float32)  {
	_ = this.PropPut(0x00000018, []interface{}{rhs})
}

func (this *Envelope) RecipientPostalfromLeft() float32 {
	retVal, _ := this.PropGet(0x00000019, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetRecipientPostalfromLeft(rhs float32)  {
	_ = this.PropPut(0x00000019, []interface{}{rhs})
}

func (this *Envelope) RecipientPostalfromTop() float32 {
	retVal, _ := this.PropGet(0x0000001a, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetRecipientPostalfromTop(rhs float32)  {
	_ = this.PropPut(0x0000001a, []interface{}{rhs})
}

func (this *Envelope) SenderNamefromLeft() float32 {
	retVal, _ := this.PropGet(0x0000001b, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetSenderNamefromLeft(rhs float32)  {
	_ = this.PropPut(0x0000001b, []interface{}{rhs})
}

func (this *Envelope) SenderNamefromTop() float32 {
	retVal, _ := this.PropGet(0x0000001c, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetSenderNamefromTop(rhs float32)  {
	_ = this.PropPut(0x0000001c, []interface{}{rhs})
}

func (this *Envelope) SenderPostalfromLeft() float32 {
	retVal, _ := this.PropGet(0x0000001d, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetSenderPostalfromLeft(rhs float32)  {
	_ = this.PropPut(0x0000001d, []interface{}{rhs})
}

func (this *Envelope) SenderPostalfromTop() float32 {
	retVal, _ := this.PropGet(0x0000001e, nil)
	return retVal.FltValVal()
}

func (this *Envelope) SetSenderPostalfromTop(rhs float32)  {
	_ = this.PropPut(0x0000001e, []interface{}{rhs})
}

var Envelope_Insert_OptArgs= []string{
	"ExtractAddress", "Address", "AutoText", "OmitReturnAddress", 
	"ReturnAddress", "ReturnAutoText", "PrintBarCode", "PrintFIMA", 
	"Size", "Height", "Width", "FeedSource", 
	"AddressFromLeft", "AddressFromTop", "ReturnAddressFromLeft", "ReturnAddressFromTop", 
	"DefaultFaceUp", "DefaultOrientation", "PrintEPostage", "Vertical", 
	"RecipientNamefromLeft", "RecipientNamefromTop", "RecipientPostalfromLeft", "RecipientPostalfromTop", 
	"SenderNamefromLeft", "SenderNamefromTop", "SenderPostalfromLeft", "SenderPostalfromTop", 
}

func (this *Envelope) Insert(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Envelope_Insert_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000069, nil, optArgs...)
	_= retVal
}

var Envelope_PrintOut_OptArgs= []string{
	"ExtractAddress", "Address", "AutoText", "OmitReturnAddress", 
	"ReturnAddress", "ReturnAutoText", "PrintBarCode", "PrintFIMA", 
	"Size", "Height", "Width", "FeedSource", 
	"AddressFromLeft", "AddressFromTop", "ReturnAddressFromLeft", "ReturnAddressFromTop", 
	"DefaultFaceUp", "DefaultOrientation", "PrintEPostage", "Vertical", 
	"RecipientNamefromLeft", "RecipientNamefromTop", "RecipientPostalfromLeft", "RecipientPostalfromTop", 
	"SenderNamefromLeft", "SenderNamefromTop", "SenderPostalfromLeft", "SenderPostalfromTop", 
}

func (this *Envelope) PrintOut(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Envelope_PrintOut_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000006a, nil, optArgs...)
	_= retVal
}

