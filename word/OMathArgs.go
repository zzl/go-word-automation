package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 8245795B-9AED-4943-A16D-E586ED8180D1
var IID_OMathArgs = syscall.GUID{0x8245795B, 0x9AED, 0x4943, 
	[8]byte{0xA1, 0x6D, 0xE5, 0x86, 0xED, 0x81, 0x80, 0xD1}}

type OMathArgs struct {
	ole.OleClient
}

func NewOMathArgs(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathArgs {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathArgs{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathArgsFromVar(v ole.Variant) *OMathArgs {
	return NewOMathArgs(v.IDispatch(), false, false)
}

func (this *OMathArgs) IID() *syscall.GUID {
	return &IID_OMathArgs
}

func (this *OMathArgs) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathArgs) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathArgs) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathArgs) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathArgs) Count() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *OMathArgs) Item(index int32) *OMath {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewOMath(retVal.IDispatch(), false, true)
}

var OMathArgs_Add_OptArgs= []string{
	"BeforeArg", 
}

func (this *OMathArgs) Add(optArgs ...interface{}) *OMath {
	optArgs = ole.ProcessOptArgs(OMathArgs_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000c8, nil, optArgs...)
	return NewOMath(retVal.IDispatch(), false, true)
}

