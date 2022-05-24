package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 5C04BD93-2F3F-4668-918D-9738EC901039
var IID_OMathRecognizedFunction = syscall.GUID{0x5C04BD93, 0x2F3F, 0x4668, 
	[8]byte{0x91, 0x8D, 0x97, 0x38, 0xEC, 0x90, 0x10, 0x39}}

type OMathRecognizedFunction struct {
	ole.OleClient
}

func NewOMathRecognizedFunction(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathRecognizedFunction {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathRecognizedFunction{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathRecognizedFunctionFromVar(v ole.Variant) *OMathRecognizedFunction {
	return NewOMathRecognizedFunction(v.IDispatch(), false, false)
}

func (this *OMathRecognizedFunction) IID() *syscall.GUID {
	return &IID_OMathRecognizedFunction
}

func (this *OMathRecognizedFunction) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathRecognizedFunction) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathRecognizedFunction) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathRecognizedFunction) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathRecognizedFunction) Index() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *OMathRecognizedFunction) Name() string {
	retVal, _ := this.PropGet(0x00000068, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *OMathRecognizedFunction) Delete()  {
	retVal, _ := this.Call(0x000000c8, nil)
	_= retVal
}

