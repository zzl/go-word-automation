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
	return NewXSLTransform(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *XSLTransform) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *XSLTransform) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *XSLTransform) Alias(allUsers bool) string {
	retVal := this.PropGet(0x00000002, []interface{}{allUsers})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XSLTransform) SetAlias(allUsers bool, rhs string)  {
	retVal := this.PropPut(0x00000002, []interface{}{allUsers, rhs})
	_= retVal
}

func (this *XSLTransform) Location(allUsers bool) string {
	retVal := this.PropGet(0x00000003, []interface{}{allUsers})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *XSLTransform) SetLocation(allUsers bool, rhs string)  {
	retVal := this.PropPut(0x00000003, []interface{}{allUsers, rhs})
	_= retVal
}

func (this *XSLTransform) Delete()  {
	retVal := this.Call(0x00000065, nil)
	_= retVal
}

func (this *XSLTransform) ID() string {
	retVal := this.PropGet(0x00000066, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

