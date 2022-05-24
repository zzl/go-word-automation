package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// AED7E08C-14F0-4F33-921D-4C5353137BF6
var IID_Editors = syscall.GUID{0xAED7E08C, 0x14F0, 0x4F33, 
	[8]byte{0x92, 0x1D, 0x4C, 0x53, 0x53, 0x13, 0x7B, 0xF6}}

type Editors struct {
	ole.OleClient
}

func NewEditors(pDisp *win32.IDispatch, addRef bool, scoped bool) *Editors {
	 if pDisp == nil {
		return nil;
	}
	p := &Editors{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func EditorsFromVar(v ole.Variant) *Editors {
	return NewEditors(v.IDispatch(), false, false)
}

func (this *Editors) IID() *syscall.GUID {
	return &IID_Editors
}

func (this *Editors) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Editors) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Editors) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Editors) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Editors) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Editors) Item(index *ole.Variant) *Editor {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewEditor(retVal.IDispatch(), false, true)
}

func (this *Editors) Add(editorID *ole.Variant) *Editor {
	retVal, _ := this.Call(0x000001f5, []interface{}{editorID})
	return NewEditor(retVal.IDispatch(), false, true)
}

