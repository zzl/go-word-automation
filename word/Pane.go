package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020960-0000-0000-C000-000000000046
var IID_Pane = syscall.GUID{0x00020960, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Pane struct {
	ole.OleClient
}

func NewPane(pDisp *win32.IDispatch, addRef bool, scoped bool) *Pane {
	p := &Pane{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PaneFromVar(v ole.Variant) *Pane {
	return NewPane(v.PdispValVal(), false, false)
}

func (this *Pane) IID() *syscall.GUID {
	return &IID_Pane
}

func (this *Pane) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Pane) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Pane) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Pane) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Pane) Document() *Document {
	retVal := this.PropGet(0x00000001, nil)
	return NewDocument(retVal.PdispValVal(), false, true)
}

func (this *Pane) Selection() *Selection {
	retVal := this.PropGet(0x00000003, nil)
	return NewSelection(retVal.PdispValVal(), false, true)
}

func (this *Pane) DisplayRulers() bool {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Pane) SetDisplayRulers(rhs bool)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *Pane) DisplayVerticalRuler() bool {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Pane) SetDisplayVerticalRuler(rhs bool)  {
	retVal := this.PropPut(0x00000005, []interface{}{rhs})
	_= retVal
}

func (this *Pane) Zooms() *Zooms {
	retVal := this.PropGet(0x00000007, nil)
	return NewZooms(retVal.PdispValVal(), false, true)
}

func (this *Pane) Index() int32 {
	retVal := this.PropGet(0x00000009, nil)
	return retVal.LValVal()
}

func (this *Pane) View() *View {
	retVal := this.PropGet(0x0000000a, nil)
	return NewView(retVal.PdispValVal(), false, true)
}

func (this *Pane) Next() *Pane {
	retVal := this.PropGet(0x0000000b, nil)
	return NewPane(retVal.PdispValVal(), false, true)
}

func (this *Pane) Previous() *Pane {
	retVal := this.PropGet(0x0000000c, nil)
	return NewPane(retVal.PdispValVal(), false, true)
}

func (this *Pane) HorizontalPercentScrolled() int32 {
	retVal := this.PropGet(0x0000000d, nil)
	return retVal.LValVal()
}

func (this *Pane) SetHorizontalPercentScrolled(rhs int32)  {
	retVal := this.PropPut(0x0000000d, []interface{}{rhs})
	_= retVal
}

func (this *Pane) VerticalPercentScrolled() int32 {
	retVal := this.PropGet(0x0000000e, nil)
	return retVal.LValVal()
}

func (this *Pane) SetVerticalPercentScrolled(rhs int32)  {
	retVal := this.PropPut(0x0000000e, []interface{}{rhs})
	_= retVal
}

func (this *Pane) MinimumFontSize() int32 {
	retVal := this.PropGet(0x0000000f, nil)
	return retVal.LValVal()
}

func (this *Pane) SetMinimumFontSize(rhs int32)  {
	retVal := this.PropPut(0x0000000f, []interface{}{rhs})
	_= retVal
}

func (this *Pane) BrowseToWindow() bool {
	retVal := this.PropGet(0x00000010, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Pane) SetBrowseToWindow(rhs bool)  {
	retVal := this.PropPut(0x00000010, []interface{}{rhs})
	_= retVal
}

func (this *Pane) BrowseWidth() int32 {
	retVal := this.PropGet(0x00000011, nil)
	return retVal.LValVal()
}

func (this *Pane) Activate()  {
	retVal := this.Call(0x00000064, nil)
	_= retVal
}

func (this *Pane) Close()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

var Pane_LargeScroll_OptArgs= []string{
	"Down", "Up", "ToRight", "ToLeft", 
}

func (this *Pane) LargeScroll(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Pane_LargeScroll_OptArgs, optArgs)
	retVal := this.Call(0x00000066, nil, optArgs...)
	_= retVal
}

var Pane_SmallScroll_OptArgs= []string{
	"Down", "Up", "ToRight", "ToLeft", 
}

func (this *Pane) SmallScroll(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Pane_SmallScroll_OptArgs, optArgs)
	retVal := this.Call(0x00000067, nil, optArgs...)
	_= retVal
}

func (this *Pane) AutoScroll(velocity int32)  {
	retVal := this.Call(0x00000068, []interface{}{velocity})
	_= retVal
}

var Pane_PageScroll_OptArgs= []string{
	"Down", "Up", 
}

func (this *Pane) PageScroll(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Pane_PageScroll_OptArgs, optArgs)
	retVal := this.Call(0x00000069, nil, optArgs...)
	_= retVal
}

func (this *Pane) NewFrameset()  {
	retVal := this.Call(0x0000006a, nil)
	_= retVal
}

func (this *Pane) TOCInFrameset()  {
	retVal := this.Call(0x0000006b, nil)
	_= retVal
}

func (this *Pane) Frameset() *Frameset {
	retVal := this.PropGet(0x00000012, nil)
	return NewFrameset(retVal.PdispValVal(), false, true)
}

func (this *Pane) Pages() *Pages {
	retVal := this.PropGet(0x00000013, nil)
	return NewPages(retVal.PdispValVal(), false, true)
}

