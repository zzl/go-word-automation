package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// CDB0FF41-E862-47BB-AE77-3FA7B1AE3189
var IID_ChartFont = syscall.GUID{0xCDB0FF41, 0xE862, 0x47BB, 
	[8]byte{0xAE, 0x77, 0x3F, 0xA7, 0xB1, 0xAE, 0x31, 0x89}}

type ChartFont struct {
	ole.OleClient
}

func NewChartFont(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartFont {
	 if pDisp == nil {
		return nil;
	}
	p := &ChartFont{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartFontFromVar(v ole.Variant) *ChartFont {
	return NewChartFont(v.IDispatch(), false, false)
}

func (this *ChartFont) IID() *syscall.GUID {
	return &IID_ChartFont
}

func (this *ChartFont) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartFont) Background() ole.Variant {
	retVal, _ := this.PropGet(0x60020000, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetBackground(rhs interface{})  {
	_ = this.PropPut(0x60020000, []interface{}{rhs})
}

func (this *ChartFont) Bold() ole.Variant {
	retVal, _ := this.PropGet(0x60020002, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetBold(rhs interface{})  {
	_ = this.PropPut(0x60020002, []interface{}{rhs})
}

func (this *ChartFont) Color() ole.Variant {
	retVal, _ := this.PropGet(0x60020004, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetColor(rhs interface{})  {
	_ = this.PropPut(0x60020004, []interface{}{rhs})
}

func (this *ChartFont) ColorIndex() ole.Variant {
	retVal, _ := this.PropGet(0x60020006, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetColorIndex(rhs interface{})  {
	_ = this.PropPut(0x60020006, []interface{}{rhs})
}

func (this *ChartFont) FontStyle() ole.Variant {
	retVal, _ := this.PropGet(0x60020008, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetFontStyle(rhs interface{})  {
	_ = this.PropPut(0x60020008, []interface{}{rhs})
}

func (this *ChartFont) Italic() ole.Variant {
	retVal, _ := this.PropGet(0x6002000a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetItalic(rhs interface{})  {
	_ = this.PropPut(0x6002000a, []interface{}{rhs})
}

func (this *ChartFont) Name() ole.Variant {
	retVal, _ := this.PropGet(0x6002000c, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetName(rhs interface{})  {
	_ = this.PropPut(0x6002000c, []interface{}{rhs})
}

func (this *ChartFont) OutlineFont() ole.Variant {
	retVal, _ := this.PropGet(0x6002000e, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetOutlineFont(rhs interface{})  {
	_ = this.PropPut(0x6002000e, []interface{}{rhs})
}

func (this *ChartFont) Shadow() ole.Variant {
	retVal, _ := this.PropGet(0x60020010, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetShadow(rhs interface{})  {
	_ = this.PropPut(0x60020010, []interface{}{rhs})
}

func (this *ChartFont) Size() ole.Variant {
	retVal, _ := this.PropGet(0x60020012, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetSize(rhs interface{})  {
	_ = this.PropPut(0x60020012, []interface{}{rhs})
}

func (this *ChartFont) StrikeThrough() ole.Variant {
	retVal, _ := this.PropGet(0x60020014, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetStrikeThrough(rhs interface{})  {
	_ = this.PropPut(0x60020014, []interface{}{rhs})
}

func (this *ChartFont) Subscript() ole.Variant {
	retVal, _ := this.PropGet(0x60020016, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetSubscript(rhs interface{})  {
	_ = this.PropPut(0x60020016, []interface{}{rhs})
}

func (this *ChartFont) Superscript() ole.Variant {
	retVal, _ := this.PropGet(0x60020018, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetSuperscript(rhs interface{})  {
	_ = this.PropPut(0x60020018, []interface{}{rhs})
}

func (this *ChartFont) Underline() ole.Variant {
	retVal, _ := this.PropGet(0x6002001a, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartFont) SetUnderline(rhs interface{})  {
	_ = this.PropPut(0x6002001a, []interface{}{rhs})
}

func (this *ChartFont) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartFont) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartFont) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

