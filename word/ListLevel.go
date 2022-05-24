package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002098D-0000-0000-C000-000000000046
var IID_ListLevel = syscall.GUID{0x0002098D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListLevel struct {
	ole.OleClient
}

func NewListLevel(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListLevel {
	 if pDisp == nil {
		return nil;
	}
	p := &ListLevel{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListLevelFromVar(v ole.Variant) *ListLevel {
	return NewListLevel(v.IDispatch(), false, false)
}

func (this *ListLevel) IID() *syscall.GUID {
	return &IID_ListLevel
}

func (this *ListLevel) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListLevel) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *ListLevel) NumberFormat() string {
	retVal, _ := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListLevel) SetNumberFormat(rhs string)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *ListLevel) TrailingCharacter() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *ListLevel) SetTrailingCharacter(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *ListLevel) NumberStyle() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *ListLevel) SetNumberStyle(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *ListLevel) NumberPosition() float32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.FltValVal()
}

func (this *ListLevel) SetNumberPosition(rhs float32)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *ListLevel) Alignment() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *ListLevel) SetAlignment(rhs int32)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *ListLevel) TextPosition() float32 {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.FltValVal()
}

func (this *ListLevel) SetTextPosition(rhs float32)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *ListLevel) TabPosition() float32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.FltValVal()
}

func (this *ListLevel) SetTabPosition(rhs float32)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *ListLevel) ResetOnHigherOld() bool {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListLevel) SetResetOnHigherOld(rhs bool)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *ListLevel) StartAt() int32 {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *ListLevel) SetStartAt(rhs int32)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *ListLevel) LinkedStyle() string {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListLevel) SetLinkedStyle(rhs string)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *ListLevel) Font() *Font {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return NewFont(retVal.IDispatch(), false, true)
}

func (this *ListLevel) SetFont(rhs *Font)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *ListLevel) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListLevel) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ListLevel) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListLevel) ResetOnHigher() int32 {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.LValVal()
}

func (this *ListLevel) SetResetOnHigher(rhs int32)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *ListLevel) PictureBullet() *InlineShape {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return NewInlineShape(retVal.IDispatch(), false, true)
}

func (this *ListLevel) ApplyPictureBullet(fileName string) *InlineShape {
	retVal, _ := this.Call(0x00000000, []interface{}{fileName})
	return NewInlineShape(retVal.IDispatch(), false, true)
}

