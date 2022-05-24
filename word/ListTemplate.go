package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002098F-0000-0000-C000-000000000046
var IID_ListTemplate = syscall.GUID{0x0002098F, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListTemplate struct {
	ole.OleClient
}

func NewListTemplate(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListTemplate {
	 if pDisp == nil {
		return nil;
	}
	p := &ListTemplate{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListTemplateFromVar(v ole.Variant) *ListTemplate {
	return NewListTemplate(v.IDispatch(), false, false)
}

func (this *ListTemplate) IID() *syscall.GUID {
	return &IID_ListTemplate
}

func (this *ListTemplate) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListTemplate) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListTemplate) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ListTemplate) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListTemplate) OutlineNumbered() bool {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListTemplate) SetOutlineNumbered(rhs bool)  {
	_ = this.PropPut(0x00000001, []interface{}{rhs})
}

func (this *ListTemplate) Name() string {
	retVal, _ := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ListTemplate) SetName(rhs string)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *ListTemplate) ListLevels() *ListLevels {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewListLevels(retVal.IDispatch(), false, true)
}

var ListTemplate_Convert_OptArgs= []string{
	"Level", 
}

func (this *ListTemplate) Convert(optArgs ...interface{}) *ListTemplate {
	optArgs = ole.ProcessOptArgs(ListTemplate_Convert_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000065, nil, optArgs...)
	return NewListTemplate(retVal.IDispatch(), false, true)
}

