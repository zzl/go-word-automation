package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209E2-0000-0000-C000-000000000046
var IID_Frameset = syscall.GUID{0x000209E2, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Frameset struct {
	ole.OleClient
}

func NewFrameset(pDisp *win32.IDispatch, addRef bool, scoped bool) *Frameset {
	p := &Frameset{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FramesetFromVar(v ole.Variant) *Frameset {
	return NewFrameset(v.PdispValVal(), false, false)
}

func (this *Frameset) IID() *syscall.GUID {
	return &IID_Frameset
}

func (this *Frameset) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Frameset) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Frameset) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Frameset) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Frameset) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Frameset) ForEach(action func(item int32) bool) {
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
		pItem, _ := v.ToInt32()
		ret := action(pItem)
		if !ret {
			break
		}
	}
}

func (this *Frameset) ParentFrameset() *Frameset {
	retVal := this.PropGet(0x000003eb, nil)
	return NewFrameset(retVal.PdispValVal(), false, true)
}

func (this *Frameset) Type() int32 {
	retVal := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *Frameset) WidthType() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Frameset) SetWidthType(rhs int32)  {
	retVal := this.PropPut(0x00000001, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) HeightType() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Frameset) SetHeightType(rhs int32)  {
	retVal := this.PropPut(0x00000002, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) Width() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Frameset) SetWidth(rhs int32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) Height() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Frameset) SetHeight(rhs int32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) ChildFramesetCount() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Frameset) ChildFramesetItem(index int32) *Frameset {
	retVal := this.PropGet(0x00000006, []interface{}{index})
	return NewFrameset(retVal.PdispValVal(), false, true)
}

func (this *Frameset) FramesetBorderWidth() float32 {
	retVal := this.PropGet(0x00000014, nil)
	return retVal.FltValVal()
}

func (this *Frameset) SetFramesetBorderWidth(rhs float32)  {
	retVal := this.PropPut(0x00000014, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) FramesetBorderColor() int32 {
	retVal := this.PropGet(0x00000015, nil)
	return retVal.LValVal()
}

func (this *Frameset) SetFramesetBorderColor(rhs int32)  {
	retVal := this.PropPut(0x00000015, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) FrameScrollbarType() int32 {
	retVal := this.PropGet(0x0000001e, nil)
	return retVal.LValVal()
}

func (this *Frameset) SetFrameScrollbarType(rhs int32)  {
	retVal := this.PropPut(0x0000001e, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) FrameResizable() bool {
	retVal := this.PropGet(0x0000001f, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Frameset) SetFrameResizable(rhs bool)  {
	retVal := this.PropPut(0x0000001f, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) FrameName() string {
	retVal := this.PropGet(0x00000022, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Frameset) SetFrameName(rhs string)  {
	retVal := this.PropPut(0x00000022, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) FrameDisplayBorders() bool {
	retVal := this.PropGet(0x00000023, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Frameset) SetFrameDisplayBorders(rhs bool)  {
	retVal := this.PropPut(0x00000023, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) FrameDefaultURL() string {
	retVal := this.PropGet(0x00000024, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Frameset) SetFrameDefaultURL(rhs string)  {
	retVal := this.PropPut(0x00000024, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) FrameLinkToFile() bool {
	retVal := this.PropGet(0x00000025, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Frameset) SetFrameLinkToFile(rhs bool)  {
	retVal := this.PropPut(0x00000025, []interface{}{rhs})
	_= retVal
}

func (this *Frameset) AddNewFrame(where int32) *Frameset {
	retVal := this.Call(0x00000032, []interface{}{where})
	return NewFrameset(retVal.PdispValVal(), false, true)
}

func (this *Frameset) Delete()  {
	retVal := this.Call(0x00000033, nil)
	_= retVal
}

