package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// E3124493-7D6A-410F-9A48-CC822C033CEC
var IID_XSLTransform = syscall.GUID{0xE3124493, 0x7D6A, 0x410F, 
	[8]byte{0x9A, 0x48, 0xCC, 0x82, 0x2C, 0x03, 0x3C, 0xEC}}

type XSLTransform struct {
	ole.OleClient
}

func NewXSLTransform(pDisp *win32.IDispatch, addRef bool, scoped bool) *XSLTransform {
	 if pDisp == nil {
		return nil;
	}
	p := &XSLTransform{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func XSLTransformFromVar(v ole.Variant) *XSLTransform {
	return NewXSLTransform(v.IDispatch(), false, false)
}

func (this *XSLTransform) IID() *syscall.GUID {
	return &IID_XSLTransform
}

func (this *XSLTransform) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *XSLTransform) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *XSLTransform) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XSLTransform) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var XSLTransform_Alias_OptArgs= []string{
	"AllUsers", 
}

func (this *XSLTransform) Alias(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(XSLTransform_Alias_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000002, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var XSLTransform_SetAlias_OptArgs= []string{
	"AllUsers", 
}

func (this *XSLTransform) SetAlias(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(XSLTransform_SetAlias_OptArgs, optArgs)
	_ = this.PropPut(0x00000002, nil, optArgs...)
}

var XSLTransform_Location_OptArgs= []string{
	"AllUsers", 
}

func (this *XSLTransform) Location(optArgs ...interface{}) string {
	optArgs = ole.ProcessOptArgs(XSLTransform_Location_OptArgs, optArgs)
	retVal, _ := this.PropGet(0x00000003, nil, optArgs...)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

var XSLTransform_SetLocation_OptArgs= []string{
	"AllUsers", 
}

func (this *XSLTransform) SetLocation(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(XSLTransform_SetLocation_OptArgs, optArgs)
	_ = this.PropPut(0x00000003, nil, optArgs...)
}

func (this *XSLTransform) Delete()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *XSLTransform) ID() string {
	retVal, _ := this.PropGet(0x00000066, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

