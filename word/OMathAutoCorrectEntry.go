package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// D8779F01-4869-4403-B334-D60C5F9C9175
var IID_OMathAutoCorrectEntry = syscall.GUID{0xD8779F01, 0x4869, 0x4403, 
	[8]byte{0xB3, 0x34, 0xD6, 0x0C, 0x5F, 0x9C, 0x91, 0x75}}

type OMathAutoCorrectEntry struct {
	ole.OleClient
}

func NewOMathAutoCorrectEntry(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathAutoCorrectEntry {
	p := &OMathAutoCorrectEntry{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathAutoCorrectEntryFromVar(v ole.Variant) *OMathAutoCorrectEntry {
	return NewOMathAutoCorrectEntry(v.PdispValVal(), false, false)
}

func (this *OMathAutoCorrectEntry) IID() *syscall.GUID {
	return &IID_OMathAutoCorrectEntry
}

func (this *OMathAutoCorrectEntry) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathAutoCorrectEntry) Application() *Application {
	retVal := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *OMathAutoCorrectEntry) Creator() int32 {
	retVal := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathAutoCorrectEntry) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *OMathAutoCorrectEntry) Index() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *OMathAutoCorrectEntry) Name() string {
	retVal := this.PropGet(0x00000068, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OMathAutoCorrectEntry) SetName(rhs string)  {
	retVal := this.PropPut(0x00000068, []interface{}{rhs})
	_= retVal
}

func (this *OMathAutoCorrectEntry) Value() string {
	retVal := this.PropGet(0x00000069, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OMathAutoCorrectEntry) SetValue(rhs string)  {
	retVal := this.PropPut(0x00000069, []interface{}{rhs})
	_= retVal
}

func (this *OMathAutoCorrectEntry) Delete()  {
	retVal := this.Call(0x000000c8, nil)
	_= retVal
}

