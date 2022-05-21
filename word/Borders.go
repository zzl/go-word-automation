package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002093C-0000-0000-C000-000000000046
var IID_Borders = syscall.GUID{0x0002093C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Borders struct {
	ole.OleClient
}

func NewBorders(pDisp *win32.IDispatch, addRef bool, scoped bool) *Borders {
	p := &Borders{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BordersFromVar(v ole.Variant) *Borders {
	return NewBorders(v.PdispValVal(), false, false)
}

func (this *Borders) IID() *syscall.GUID {
	return &IID_Borders
}

func (this *Borders) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Borders) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Borders) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Borders) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Borders) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Borders) ForEach(action func(item *Border) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*Border)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Borders) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Borders) Enable() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Borders) SetEnable(rhs int32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *Borders) DistanceFromTop() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Borders) SetDistanceFromTop(rhs int32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *Borders) Shadow() bool {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Borders) SetShadow(rhs bool)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *Borders) InsideLineStyle() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Borders) SetInsideLineStyle(rhs int32)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *Borders) OutsideLineStyle() int32 {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *Borders) SetOutsideLineStyle(rhs int32)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *Borders) InsideLineWidth() int32 {
	retVal := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *Borders) SetInsideLineWidth(rhs int32)  {
	retVal := this.PropPut(0x00000008, []interface{}{rhs})
	_= retVal
}

func (this *Borders) OutsideLineWidth() int32 {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.LValVal()
}

func (this *Borders) SetOutsideLineWidth(rhs int32)  {
	retVal := this.PropPut(0x00000009, []interface{}{rhs})
	_= retVal
}

func (this *Borders) InsideColorIndex() int32 {
	retVal := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *Borders) SetInsideColorIndex(rhs int32)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *Borders) OutsideColorIndex() int32 {
	retVal := this.PropGet(0x0000000b, nil)
	return retVal.LValVal()
}

func (this *Borders) SetOutsideColorIndex(rhs int32)  {
	retVal := this.PropPut(0x0000000b, []interface{}{rhs})
	_= retVal
}

func (this *Borders) DistanceFromLeft() int32 {
	retVal := this.PropGet(0x00000014, nil)
	return retVal.LValVal()
}

func (this *Borders) SetDistanceFromLeft(rhs int32)  {
	retVal := this.PropPut(0x00000014, []interface{}{rhs})
	_= retVal
}

func (this *Borders) DistanceFromBottom() int32 {
	retVal := this.PropGet(0x00000015, nil)
	return retVal.LValVal()
}

func (this *Borders) SetDistanceFromBottom(rhs int32)  {
	retVal := this.PropPut(0x00000015, []interface{}{rhs})
	_= retVal
}

func (this *Borders) DistanceFromRight() int32 {
	retVal := this.PropGet(0x00000016, nil)
	return retVal.LValVal()
}

func (this *Borders) SetDistanceFromRight(rhs int32)  {
	retVal := this.PropPut(0x00000016, []interface{}{rhs})
	_= retVal
}

func (this *Borders) AlwaysInFront() bool {
	retVal := this.PropGet(0x00000017, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Borders) SetAlwaysInFront(rhs bool)  {
	retVal := this.PropPut(0x00000017, []interface{}{rhs})
	_= retVal
}

func (this *Borders) SurroundHeader() bool {
	retVal := this.PropGet(0x00000018, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Borders) SetSurroundHeader(rhs bool)  {
	retVal := this.PropPut(0x00000018, []interface{}{rhs})
	_= retVal
}

func (this *Borders) SurroundFooter() bool {
	retVal := this.PropGet(0x00000019, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Borders) SetSurroundFooter(rhs bool)  {
	retVal := this.PropPut(0x00000019, []interface{}{rhs})
	_= retVal
}

func (this *Borders) JoinBorders() bool {
	retVal := this.PropGet(0x0000001a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Borders) SetJoinBorders(rhs bool)  {
	retVal := this.PropPut(0x0000001a, []interface{}{rhs})
	_= retVal
}

func (this *Borders) HasHorizontal() bool {
	retVal := this.PropGet(0x0000001b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Borders) HasVertical() bool {
	retVal := this.PropGet(0x0000001c, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Borders) DistanceFrom() int32 {
	retVal := this.PropGet(0x0000001d, nil)
	return retVal.LValVal()
}

func (this *Borders) SetDistanceFrom(rhs int32)  {
	retVal := this.PropPut(0x0000001d, []interface{}{rhs})
	_= retVal
}

func (this *Borders) EnableFirstPageInSection() bool {
	retVal := this.PropGet(0x0000001e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Borders) SetEnableFirstPageInSection(rhs bool)  {
	retVal := this.PropPut(0x0000001e, []interface{}{rhs})
	_= retVal
}

func (this *Borders) EnableOtherPagesInSection() bool {
	retVal := this.PropGet(0x0000001f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Borders) SetEnableOtherPagesInSection(rhs bool)  {
	retVal := this.PropPut(0x0000001f, []interface{}{rhs})
	_= retVal
}

func (this *Borders) Item(index int32) *Border {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewBorder(retVal.PdispValVal(), false, true)
}

func (this *Borders) ApplyPageBordersToAllSections()  {
	retVal := this.Call(0x000007d0, nil)
	_= retVal
}

func (this *Borders) InsideColor() int32 {
	retVal := this.PropGet(0x00000020, nil)
	return retVal.LValVal()
}

func (this *Borders) SetInsideColor(rhs int32)  {
	retVal := this.PropPut(0x00000020, []interface{}{rhs})
	_= retVal
}

func (this *Borders) OutsideColor() int32 {
	retVal := this.PropGet(0x00000021, nil)
	return retVal.LValVal()
}

func (this *Borders) SetOutsideColor(rhs int32)  {
	retVal := this.PropPut(0x00000021, []interface{}{rhs})
	_= retVal
}

