package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// E59544D5-C299-46A0-84C1-C51AB38F9759
var IID_CoAuthor = syscall.GUID{0xE59544D5, 0xC299, 0x46A0, 
	[8]byte{0x84, 0xC1, 0xC5, 0x1A, 0xB3, 0x8F, 0x97, 0x59}}

type CoAuthor struct {
	ole.OleClient
}

func NewCoAuthor(pDisp *win32.IDispatch, addRef bool, scoped bool) *CoAuthor {
	 if pDisp == nil {
		return nil;
	}
	p := &CoAuthor{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CoAuthorFromVar(v ole.Variant) *CoAuthor {
	return NewCoAuthor(v.IDispatch(), false, false)
}

func (this *CoAuthor) IID() *syscall.GUID {
	return &IID_CoAuthor
}

func (this *CoAuthor) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *CoAuthor) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *CoAuthor) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *CoAuthor) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *CoAuthor) ID() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CoAuthor) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *CoAuthor) IsMe() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *CoAuthor) Locks() *CoAuthLocks {
	retVal, _ := this.PropGet(0x00000004, nil)
	return NewCoAuthLocks(retVal.IDispatch(), false, true)
}

func (this *CoAuthor) EmailAddress() string {
	retVal, _ := this.PropGet(0x00000005, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

