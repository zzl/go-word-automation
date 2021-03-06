package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209E3-0000-0000-C000-000000000046
var IID_DefaultWebOptions = syscall.GUID{0x000209E3, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DefaultWebOptions struct {
	ole.OleClient
}

func NewDefaultWebOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *DefaultWebOptions {
	 if pDisp == nil {
		return nil;
	}
	p := &DefaultWebOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DefaultWebOptionsFromVar(v ole.Variant) *DefaultWebOptions {
	return NewDefaultWebOptions(v.IDispatch(), false, false)
}

func (this *DefaultWebOptions) IID() *syscall.GUID {
	return &IID_DefaultWebOptions
}

func (this *DefaultWebOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DefaultWebOptions) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *DefaultWebOptions) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DefaultWebOptions) OptimizeForBrowser() bool {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetOptimizeForBrowser(rhs bool)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *DefaultWebOptions) BrowserLevel() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetBrowserLevel(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *DefaultWebOptions) RelyOnCSS() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetRelyOnCSS(rhs bool)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *DefaultWebOptions) OrganizeInFolder() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetOrganizeInFolder(rhs bool)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *DefaultWebOptions) UpdateLinksOnSave() bool {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetUpdateLinksOnSave(rhs bool)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *DefaultWebOptions) UseLongFileNames() bool {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetUseLongFileNames(rhs bool)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *DefaultWebOptions) CheckIfOfficeIsHTMLEditor() bool {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetCheckIfOfficeIsHTMLEditor(rhs bool)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *DefaultWebOptions) CheckIfWordIsDefaultHTMLEditor() bool {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetCheckIfWordIsDefaultHTMLEditor(rhs bool)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *DefaultWebOptions) RelyOnVML() bool {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetRelyOnVML(rhs bool)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *DefaultWebOptions) AllowPNG() bool {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetAllowPNG(rhs bool)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *DefaultWebOptions) ScreenSize() int32 {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetScreenSize(rhs int32)  {
	_ = this.PropPut(0x0000000b, []interface{}{rhs})
}

func (this *DefaultWebOptions) PixelsPerInch() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetPixelsPerInch(rhs int32)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

func (this *DefaultWebOptions) Encoding() int32 {
	retVal, _ := this.PropGet(0x0000000d, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetEncoding(rhs int32)  {
	_ = this.PropPut(0x0000000d, []interface{}{rhs})
}

func (this *DefaultWebOptions) AlwaysSaveInDefaultEncoding() bool {
	retVal, _ := this.PropGet(0x0000000e, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetAlwaysSaveInDefaultEncoding(rhs bool)  {
	_ = this.PropPut(0x0000000e, []interface{}{rhs})
}

func (this *DefaultWebOptions) Fonts() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x0000000f, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *DefaultWebOptions) FolderSuffix() string {
	retVal, _ := this.PropGet(0x00000010, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DefaultWebOptions) TargetBrowser() int32 {
	retVal, _ := this.PropGet(0x00000011, nil)
	return retVal.LValVal()
}

func (this *DefaultWebOptions) SetTargetBrowser(rhs int32)  {
	_ = this.PropPut(0x00000011, []interface{}{rhs})
}

func (this *DefaultWebOptions) SaveNewWebPagesAsWebArchives() bool {
	retVal, _ := this.PropGet(0x00000012, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *DefaultWebOptions) SetSaveNewWebPagesAsWebArchives(rhs bool)  {
	_ = this.PropPut(0x00000012, []interface{}{rhs})
}

