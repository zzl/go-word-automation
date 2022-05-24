package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// E2E0F3A7-204C-40C5-BAA5-290F374FDF5A
var IID_OMathBreaks = syscall.GUID{0xE2E0F3A7, 0x204C, 0x40C5, 
	[8]byte{0xBA, 0xA5, 0x29, 0x0F, 0x37, 0x4F, 0xDF, 0x5A}}

type OMathBreaks struct {
	ole.OleClient
}

func NewOMathBreaks(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathBreaks {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathBreaks{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathBreaksFromVar(v ole.Variant) *OMathBreaks {
	return NewOMathBreaks(v.IDispatch(), false, false)
}

func (this *OMathBreaks) IID() *syscall.GUID {
	return &IID_OMathBreaks
}

func (this *OMathBreaks) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathBreaks) Application() *Application {
	retVal, _ := this.PropGet(0x00000065, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathBreaks) Creator() int32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *OMathBreaks) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000067, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathBreaks) Count() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *OMathBreaks) Item(index int32) *OMathBreak {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewOMathBreak(retVal.IDispatch(), false, true)
}

func (this *OMathBreaks) Add(range_ *Range) *OMathBreak {
	retVal, _ := this.Call(0x00000069, []interface{}{range_})
	return NewOMathBreak(retVal.IDispatch(), false, true)
}

