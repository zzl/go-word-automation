package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002099D-0000-0000-C000-000000000046
var IID_Hyperlink = syscall.GUID{0x0002099D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Hyperlink struct {
	ole.OleClient
}

func NewHyperlink(pDisp *win32.IDispatch, addRef bool, scoped bool) *Hyperlink {
	 if pDisp == nil {
		return nil;
	}
	p := &Hyperlink{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HyperlinkFromVar(v ole.Variant) *Hyperlink {
	return NewHyperlink(v.IDispatch(), false, false)
}

func (this *Hyperlink) IID() *syscall.GUID {
	return &IID_Hyperlink
}

func (this *Hyperlink) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Hyperlink) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Hyperlink) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Hyperlink) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Hyperlink) Name() string {
	retVal, _ := this.PropGet(0x000003eb, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) AddressOld() string {
	retVal, _ := this.PropGet(0x000003ec, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) Type() int32 {
	retVal, _ := this.PropGet(0x000003ed, nil)
	return retVal.LValVal()
}

func (this *Hyperlink) Range() *Range {
	retVal, _ := this.PropGet(0x000003ee, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Hyperlink) Shape() *Shape {
	retVal, _ := this.PropGet(0x000003ef, nil)
	return NewShape(retVal.IDispatch(), false, true)
}

func (this *Hyperlink) SubAddressOld() string {
	retVal, _ := this.PropGet(0x000003f0, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) ExtraInfoRequired() bool {
	retVal, _ := this.PropGet(0x000003f1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Hyperlink) Delete()  {
	retVal, _ := this.Call(0x00000067, nil)
	_= retVal
}

var Hyperlink_Follow_OptArgs= []string{
	"NewWindow", "AddHistory", "ExtraInfo", "Method", "HeaderInfo", 
}

func (this *Hyperlink) Follow(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Hyperlink_Follow_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000068, nil, optArgs...)
	_= retVal
}

func (this *Hyperlink) AddToFavorites()  {
	retVal, _ := this.Call(0x00000069, nil)
	_= retVal
}

func (this *Hyperlink) CreateNewDocument(fileName string, editNow bool, overwrite bool)  {
	retVal, _ := this.Call(0x0000006a, []interface{}{fileName, editNow, overwrite})
	_= retVal
}

func (this *Hyperlink) Address() string {
	retVal, _ := this.PropGet(0x0000044c, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetAddress(rhs string)  {
	_ = this.PropPut(0x0000044c, []interface{}{rhs})
}

func (this *Hyperlink) SubAddress() string {
	retVal, _ := this.PropGet(0x0000044d, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetSubAddress(rhs string)  {
	_ = this.PropPut(0x0000044d, []interface{}{rhs})
}

func (this *Hyperlink) EmailSubject() string {
	retVal, _ := this.PropGet(0x000003f2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetEmailSubject(rhs string)  {
	_ = this.PropPut(0x000003f2, []interface{}{rhs})
}

func (this *Hyperlink) ScreenTip() string {
	retVal, _ := this.PropGet(0x000003f3, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetScreenTip(rhs string)  {
	_ = this.PropPut(0x000003f3, []interface{}{rhs})
}

func (this *Hyperlink) TextToDisplay() string {
	retVal, _ := this.PropGet(0x000003f4, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetTextToDisplay(rhs string)  {
	_ = this.PropPut(0x000003f4, []interface{}{rhs})
}

func (this *Hyperlink) Target() string {
	retVal, _ := this.PropGet(0x000003f5, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Hyperlink) SetTarget(rhs string)  {
	_ = this.PropPut(0x000003f5, []interface{}{rhs})
}

