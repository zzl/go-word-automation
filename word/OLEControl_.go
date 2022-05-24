package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209A4-0000-0000-C000-000000000046
var IID_OLEControl_ = syscall.GUID{0x000209A4, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type OLEControl_ struct {
	ole.OleClient
}

func NewOLEControl_(pDisp *win32.IDispatch, addRef bool, scoped bool) *OLEControl_ {
	 if pDisp == nil {
		return nil;
	}
	p := &OLEControl_{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OLEControl_FromVar(v ole.Variant) *OLEControl_ {
	return NewOLEControl_(v.IDispatch(), false, false)
}

func (this *OLEControl_) IID() *syscall.GUID {
	return &IID_OLEControl_
}

func (this *OLEControl_) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OLEControl_) Left() float32 {
	retVal, _ := this.PropGet(-2147417853, nil)
	return retVal.FltValVal()
}

func (this *OLEControl_) SetLeft(rhs float32)  {
	_ = this.PropPut(-2147417853, []interface{}{rhs})
}

func (this *OLEControl_) Top() float32 {
	retVal, _ := this.PropGet(-2147417852, nil)
	return retVal.FltValVal()
}

func (this *OLEControl_) SetTop(rhs float32)  {
	_ = this.PropPut(-2147417852, []interface{}{rhs})
}

func (this *OLEControl_) Height() float32 {
	retVal, _ := this.PropGet(-2147417851, nil)
	return retVal.FltValVal()
}

func (this *OLEControl_) SetHeight(rhs float32)  {
	_ = this.PropPut(-2147417851, []interface{}{rhs})
}

func (this *OLEControl_) Width() float32 {
	retVal, _ := this.PropGet(-2147417850, nil)
	return retVal.FltValVal()
}

func (this *OLEControl_) SetWidth(rhs float32)  {
	_ = this.PropPut(-2147417850, []interface{}{rhs})
}

func (this *OLEControl_) Name() string {
	retVal, _ := this.PropGet(-2147418112, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEControl_) SetName(rhs string)  {
	_ = this.PropPut(-2147418112, []interface{}{rhs})
}

func (this *OLEControl_) Automation() *ole.DispatchClass {
	retVal, _ := this.PropGet(-2147417849, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OLEControl_) Select()  {
	retVal, _ := this.Call(-2147417568, nil)
	_= retVal
}

func (this *OLEControl_) Copy()  {
	retVal, _ := this.Call(-2147417560, nil)
	_= retVal
}

func (this *OLEControl_) Cut()  {
	retVal, _ := this.Call(-2147417559, nil)
	_= retVal
}

func (this *OLEControl_) Delete()  {
	retVal, _ := this.Call(-2147417520, nil)
	_= retVal
}

func (this *OLEControl_) Activate()  {
	retVal, _ := this.Call(-2147417519, nil)
	_= retVal
}

func (this *OLEControl_) AltHTML() string {
	retVal, _ := this.PropGet(-2147415101, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OLEControl_) SetAltHTML(rhs string)  {
	_ = this.PropPut(-2147415101, []interface{}{rhs})
}

