package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020939-0000-0000-C000-000000000046
var IID_TextRetrievalMode = syscall.GUID{0x00020939, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TextRetrievalMode struct {
	ole.OleClient
}

func NewTextRetrievalMode(pDisp *win32.IDispatch, addRef bool, scoped bool) *TextRetrievalMode {
	 if pDisp == nil {
		return nil;
	}
	p := &TextRetrievalMode{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TextRetrievalModeFromVar(v ole.Variant) *TextRetrievalMode {
	return NewTextRetrievalMode(v.IDispatch(), false, false)
}

func (this *TextRetrievalMode) IID() *syscall.GUID {
	return &IID_TextRetrievalMode
}

func (this *TextRetrievalMode) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TextRetrievalMode) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TextRetrievalMode) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TextRetrievalMode) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextRetrievalMode) ViewType() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *TextRetrievalMode) SetViewType(rhs int32)  {
	_ = this.PropPut(0x00000000, []interface{}{rhs})
}

func (this *TextRetrievalMode) Duplicate() *TextRetrievalMode {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewTextRetrievalMode(retVal.IDispatch(), false, true)
}

func (this *TextRetrievalMode) IncludeHiddenText() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextRetrievalMode) SetIncludeHiddenText(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *TextRetrievalMode) IncludeFieldCodes() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *TextRetrievalMode) SetIncludeFieldCodes(rhs bool)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

