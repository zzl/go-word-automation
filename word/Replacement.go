package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209B1-0000-0000-C000-000000000046
var IID_Replacement = syscall.GUID{0x000209B1, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Replacement struct {
	ole.OleClient
}

func NewReplacement(pDisp *win32.IDispatch, addRef bool, scoped bool) *Replacement {
	p := &Replacement{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ReplacementFromVar(v ole.Variant) *Replacement {
	return NewReplacement(v.PdispValVal(), false, false)
}

func (this *Replacement) IID() *syscall.GUID {
	return &IID_Replacement
}

func (this *Replacement) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Replacement) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Replacement) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Replacement) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Replacement) Font() *Font {
	retVal := this.PropGet(0x0000000a, nil)
	return NewFont(retVal.PdispValVal(), false, true)
}

func (this *Replacement) SetFont(rhs *Font)  {
	retVal := this.PropPut(0x0000000a, []interface{}{rhs})
	_= retVal
}

func (this *Replacement) ParagraphFormat() *ParagraphFormat {
	retVal := this.PropGet(0x0000000b, nil)
	return NewParagraphFormat(retVal.PdispValVal(), false, true)
}

func (this *Replacement) SetParagraphFormat(rhs *ParagraphFormat)  {
	retVal := this.PropPut(0x0000000b, []interface{}{rhs})
	_= retVal
}

func (this *Replacement) Style() ole.Variant {
	retVal := this.PropGet(0x0000000c, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Replacement) SetStyle(rhs *ole.Variant)  {
	retVal := this.PropPut(0x0000000c, []interface{}{rhs})
	_= retVal
}

func (this *Replacement) Text() string {
	retVal := this.PropGet(0x0000000f, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Replacement) SetText(rhs string)  {
	retVal := this.PropPut(0x0000000f, []interface{}{rhs})
	_= retVal
}

func (this *Replacement) LanguageID() int32 {
	retVal := this.PropGet(0x00000010, nil)
	return retVal.LValVal()
}

func (this *Replacement) SetLanguageID(rhs int32)  {
	retVal := this.PropPut(0x00000010, []interface{}{rhs})
	_= retVal
}

func (this *Replacement) Highlight() int32 {
	retVal := this.PropGet(0x00000011, nil)
	return retVal.LValVal()
}

func (this *Replacement) SetHighlight(rhs int32)  {
	retVal := this.PropPut(0x00000011, []interface{}{rhs})
	_= retVal
}

func (this *Replacement) Frame() *Frame {
	retVal := this.PropGet(0x00000012, nil)
	return NewFrame(retVal.PdispValVal(), false, true)
}

func (this *Replacement) LanguageIDFarEast() int32 {
	retVal := this.PropGet(0x00000013, nil)
	return retVal.LValVal()
}

func (this *Replacement) SetLanguageIDFarEast(rhs int32)  {
	retVal := this.PropPut(0x00000013, []interface{}{rhs})
	_= retVal
}

func (this *Replacement) ClearFormatting()  {
	retVal := this.Call(0x00000014, nil)
	_= retVal
}

func (this *Replacement) NoProofing() int32 {
	retVal := this.PropGet(0x00000015, nil)
	return retVal.LValVal()
}

func (this *Replacement) SetNoProofing(rhs int32)  {
	retVal := this.PropPut(0x00000015, []interface{}{rhs})
	_= retVal
}

