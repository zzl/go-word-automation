package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020979-0000-0000-C000-000000000046
var IID_CaptionLabel = syscall.GUID{0x00020979, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type CaptionLabel struct {
	ole.OleClient
}

func NewCaptionLabel(pDisp *win32.IDispatch, addRef bool, scoped bool) *CaptionLabel {
	p := &CaptionLabel{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CaptionLabelFromVar(v ole.Variant) *CaptionLabel {
	return NewCaptionLabel(v.PdispValVal(), false, false)
}

func (this *CaptionLabel) IID() *syscall.GUID {
	return &IID_CaptionLabel
}

func (this *CaptionLabel) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CaptionLabel) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *CaptionLabel) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CaptionLabel) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *CaptionLabel) Name() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CaptionLabel) BuiltIn() bool {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CaptionLabel) ID() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *CaptionLabel) IncludeChapterNumber() bool {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CaptionLabel) SetIncludeChapterNumber(rhs bool)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *CaptionLabel) NumberStyle() int32 {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *CaptionLabel) SetNumberStyle(rhs int32)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *CaptionLabel) ChapterStyleLevel() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *CaptionLabel) SetChapterStyleLevel(rhs int32)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *CaptionLabel) Separator() int32 {
	retVal := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *CaptionLabel) SetSeparator(rhs int32)  {
	retVal := this.PropPut(0x00000006, []interface{}{rhs})
	_= retVal
}

func (this *CaptionLabel) Position() int32 {
	retVal := this.PropGet(0x00000007, nil)
	return retVal.LValVal()
}

func (this *CaptionLabel) SetPosition(rhs int32)  {
	retVal := this.PropPut(0x00000007, []interface{}{rhs})
	_= retVal
}

func (this *CaptionLabel) Delete()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

