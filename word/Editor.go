package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// DD947D72-F33C-4198-9BDF-F86181D05E41
var IID_Editor = syscall.GUID{0xDD947D72, 0xF33C, 0x4198, 
	[8]byte{0x9B, 0xDF, 0xF8, 0x61, 0x81, 0xD0, 0x5E, 0x41}}

type Editor struct {
	ole.OleClient
}

func NewEditor(pDisp *win32.IDispatch, addRef bool, scoped bool) *Editor {
	 if pDisp == nil {
		return nil;
	}
	p := &Editor{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EditorFromVar(v ole.Variant) *Editor {
	return NewEditor(v.IDispatch(), false, false)
}

func (this *Editor) IID() *syscall.GUID {
	return &IID_Editor
}

func (this *Editor) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Editor) ID() string {
	retVal, _ := this.PropGet(0x00000064, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Editor) Name() string {
	retVal, _ := this.PropGet(0x00000065, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Editor) Range() *Range {
	retVal, _ := this.PropGet(0x00000066, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Editor) NextRange() *Range {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Editor) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Editor) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Editor) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Editor) Delete()  {
	retVal, _ := this.Call(0x000001f4, nil)
	_= retVal
}

func (this *Editor) DeleteAll()  {
	retVal, _ := this.Call(0x000001f5, nil)
	_= retVal
}

func (this *Editor) SelectAll()  {
	retVal, _ := this.Call(0x000001f6, nil)
	_= retVal
}

