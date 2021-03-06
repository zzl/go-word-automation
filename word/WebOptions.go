package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209E4-0000-0000-C000-000000000046
var IID_WebOptions = syscall.GUID{0x000209E4, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type WebOptions struct {
	ole.OleClient
}

func NewWebOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *WebOptions {
	 if pDisp == nil {
		return nil;
	}
	p := &WebOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func WebOptionsFromVar(v ole.Variant) *WebOptions {
	return NewWebOptions(v.IDispatch(), false, false)
}

func (this *WebOptions) IID() *syscall.GUID {
	return &IID_WebOptions
}

func (this *WebOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *WebOptions) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *WebOptions) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *WebOptions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *WebOptions) OptimizeForBrowser() bool {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetOptimizeForBrowser(rhs bool)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *WebOptions) BrowserLevel() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *WebOptions) SetBrowserLevel(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *WebOptions) RelyOnCSS() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetRelyOnCSS(rhs bool)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *WebOptions) OrganizeInFolder() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetOrganizeInFolder(rhs bool)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *WebOptions) UseLongFileNames() bool {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetUseLongFileNames(rhs bool)  {
	_ = this.PropPut(0x00000005, []interface{}{rhs})
}

func (this *WebOptions) RelyOnVML() bool {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetRelyOnVML(rhs bool)  {
	_ = this.PropPut(0x00000006, []interface{}{rhs})
}

func (this *WebOptions) AllowPNG() bool {
	retVal, _ := this.PropGet(0x00000007, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *WebOptions) SetAllowPNG(rhs bool)  {
	_ = this.PropPut(0x00000007, []interface{}{rhs})
}

func (this *WebOptions) ScreenSize() int32 {
	retVal, _ := this.PropGet(0x00000008, nil)
	return retVal.LValVal()
}

func (this *WebOptions) SetScreenSize(rhs int32)  {
	_ = this.PropPut(0x00000008, []interface{}{rhs})
}

func (this *WebOptions) PixelsPerInch() int32 {
	retVal, _ := this.PropGet(0x00000009, nil)
	return retVal.LValVal()
}

func (this *WebOptions) SetPixelsPerInch(rhs int32)  {
	_ = this.PropPut(0x00000009, []interface{}{rhs})
}

func (this *WebOptions) Encoding() int32 {
	retVal, _ := this.PropGet(0x0000000a, nil)
	return retVal.LValVal()
}

func (this *WebOptions) SetEncoding(rhs int32)  {
	_ = this.PropPut(0x0000000a, []interface{}{rhs})
}

func (this *WebOptions) FolderSuffix() string {
	retVal, _ := this.PropGet(0x0000000b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *WebOptions) UseDefaultFolderSuffix()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *WebOptions) TargetBrowser() int32 {
	retVal, _ := this.PropGet(0x0000000c, nil)
	return retVal.LValVal()
}

func (this *WebOptions) SetTargetBrowser(rhs int32)  {
	_ = this.PropPut(0x0000000c, []interface{}{rhs})
}

