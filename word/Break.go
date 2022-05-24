package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 79635BF1-BD1D-4B3F-A520-C1106F1AAAD8
var IID_Break = syscall.GUID{0x79635BF1, 0xBD1D, 0x4B3F, 
	[8]byte{0xA5, 0x20, 0xC1, 0x10, 0x6F, 0x1A, 0xAA, 0xD8}}

type Break struct {
	ole.OleClient
}

func NewBreak(pDisp *win32.IDispatch, addRef bool, scoped bool) *Break {
	 if pDisp == nil {
		return nil;
	}
	p := &Break{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BreakFromVar(v ole.Variant) *Break {
	return NewBreak(v.IDispatch(), false, false)
}

func (this *Break) IID() *syscall.GUID {
	return &IID_Break
}

func (this *Break) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Break) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Break) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Break) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Break) Range() *Range {
	retVal, _ := this.PropGet(0x00000002, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Break) PageIndex() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

