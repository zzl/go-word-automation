package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209ED-0000-0000-C000-000000000046
var IID_SmartTag = syscall.GUID{0x000209ED, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SmartTag struct {
	ole.OleClient
}

func NewSmartTag(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTag {
	p := &SmartTag{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagFromVar(v ole.Variant) *SmartTag {
	return NewSmartTag(v.PdispValVal(), false, false)
}

func (this *SmartTag) IID() *syscall.GUID {
	return &IID_SmartTag
}

func (this *SmartTag) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTag) Name() string {
	retVal := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTag) XML() string {
	retVal := this.PropGet(0x00000002, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTag) Range() *Range {
	retVal := this.PropGet(0x00000003, nil)
	return NewRange(retVal.PdispValVal(), false, true)
}

func (this *SmartTag) DownloadURL() string {
	retVal := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SmartTag) Properties() *CustomProperties {
	retVal := this.PropGet(0x00000005, nil)
	return NewCustomProperties(retVal.PdispValVal(), false, true)
}

func (this *SmartTag) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *SmartTag) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SmartTag) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SmartTag) Select()  {
	retVal := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *SmartTag) Delete()  {
	retVal := this.Call(0x0000000b, nil)
	_= retVal
}

func (this *SmartTag) SmartTagActions() *SmartTagActions {
	retVal := this.PropGet(0x000003eb, nil)
	return NewSmartTagActions(retVal.PdispValVal(), false, true)
}

func (this *SmartTag) XMLNode() *XMLNode {
	retVal := this.PropGet(0x000003ec, nil)
	return NewXMLNode(retVal.PdispValVal(), false, true)
}

