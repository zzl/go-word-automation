package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020972-0000-0000-C000-000000000046
var IID_LineNumbering = syscall.GUID{0x00020972, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type LineNumbering struct {
	ole.OleClient
}

func NewLineNumbering(pDisp *win32.IDispatch, addRef bool, scoped bool) *LineNumbering {
	p := &LineNumbering{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LineNumberingFromVar(v ole.Variant) *LineNumbering {
	return NewLineNumbering(v.PdispValVal(), false, false)
}

func (this *LineNumbering) IID() *syscall.GUID {
	return &IID_LineNumbering
}

func (this *LineNumbering) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *LineNumbering) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *LineNumbering) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *LineNumbering) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *LineNumbering) RestartMode() int32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *LineNumbering) SetRestartMode(rhs int32)  {
	retVal := this.PropPut(0x00000064, []interface{}{rhs})
	_= retVal
}

func (this *LineNumbering) StartingNumber() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *LineNumbering) SetStartingNumber(rhs int32)  {
	retVal := this.PropPut(0x00000065, []interface{}{rhs})
	_= retVal
}

func (this *LineNumbering) DistanceFromText() float32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *LineNumbering) SetDistanceFromText(rhs float32)  {
	retVal := this.PropPut(0x00000066, []interface{}{rhs})
	_= retVal
}

func (this *LineNumbering) CountBy() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *LineNumbering) SetCountBy(rhs int32)  {
	retVal := this.PropPut(0x00000067, []interface{}{rhs})
	_= retVal
}

func (this *LineNumbering) Active() int32 {
	retVal := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *LineNumbering) SetActive(rhs int32)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

