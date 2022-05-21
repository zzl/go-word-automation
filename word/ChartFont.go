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
	return NewChartFont(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x60020000, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetBackground(rhs interface{})  {
	retVal := this.PropPut(0x60020000, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Bold() ole.Variant {
	retVal := this.PropGet(0x60020002, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetBold(rhs interface{})  {
	retVal := this.PropPut(0x60020002, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Color() ole.Variant {
	retVal := this.PropGet(0x60020004, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetColor(rhs interface{})  {
	retVal := this.PropPut(0x60020004, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) ColorIndex() ole.Variant {
	retVal := this.PropGet(0x60020006, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetColorIndex(rhs interface{})  {
	retVal := this.PropPut(0x60020006, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) FontStyle() ole.Variant {
	retVal := this.PropGet(0x60020008, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetFontStyle(rhs interface{})  {
	retVal := this.PropPut(0x60020008, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Italic() ole.Variant {
	retVal := this.PropGet(0x6002000a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetItalic(rhs interface{})  {
	retVal := this.PropPut(0x6002000a, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Name() ole.Variant {
	retVal := this.PropGet(0x6002000c, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetName(rhs interface{})  {
	retVal := this.PropPut(0x6002000c, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) OutlineFont() ole.Variant {
	retVal := this.PropGet(0x6002000e, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetOutlineFont(rhs interface{})  {
	retVal := this.PropPut(0x6002000e, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Shadow() ole.Variant {
	retVal := this.PropGet(0x60020010, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetShadow(rhs interface{})  {
	retVal := this.PropPut(0x60020010, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Size() ole.Variant {
	retVal := this.PropGet(0x60020012, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetSize(rhs interface{})  {
	retVal := this.PropPut(0x60020012, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) StrikeThrough() ole.Variant {
	retVal := this.PropGet(0x60020014, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetStrikeThrough(rhs interface{})  {
	retVal := this.PropPut(0x60020014, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Subscript() ole.Variant {
	retVal := this.PropGet(0x60020016, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetSubscript(rhs interface{})  {
	retVal := this.PropPut(0x60020016, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Superscript() ole.Variant {
	retVal := this.PropGet(0x60020018, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetSuperscript(rhs interface{})  {
	retVal := this.PropPut(0x60020018, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Underline() ole.Variant {
	retVal := this.PropGet(0x6002001a, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *ChartFont) SetUnderline(rhs interface{})  {
	retVal := this.PropPut(0x6002001a, []interface{}{rhs})
	_= retVal
}

func (this *ChartFont) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartFont) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *ChartFont) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

