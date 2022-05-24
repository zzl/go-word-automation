package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// F1F37152-1DB1-4901-AD9A-C740F99464B4
var IID_OMathFunction = syscall.GUID{0xF1F37152, 0x1DB1, 0x4901, 
	[8]byte{0xAD, 0x9A, 0xC7, 0x40, 0xF9, 0x94, 0x64, 0xB4}}

type OMathFunction struct {
	ole.OleClient
}

func NewOMathFunction(pDisp *win32.IDispatch, addRef bool, scoped bool) *OMathFunction {
	 if pDisp == nil {
		return nil;
	}
	p := &OMathFunction{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func OMathFunctionFromVar(v ole.Variant) *OMathFunction {
	return NewOMathFunction(v.IDispatch(), false, false)
}

func (this *OMathFunction) IID() *syscall.GUID {
	return &IID_OMathFunction
}

func (this *OMathFunction) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *OMathFunction) Type() int32 {
	retVal, _ := this.PropGet(0x00000000, nil)
	return retVal.LValVal()
}

func (this *OMathFunction) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *OMathFunction) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *OMathFunction) Range() *Range {
	retVal, _ := this.PropGet(0x00000067, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Args() *OMathArgs {
	retVal, _ := this.PropGet(0x00000068, nil)
	return NewOMathArgs(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Acc() *OMathAcc {
	retVal, _ := this.PropGet(0x00000069, nil)
	return NewOMathAcc(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Bar() *OMathBar {
	retVal, _ := this.PropGet(0x0000006a, nil)
	return NewOMathBar(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Box() *OMathBox {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return NewOMathBox(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) BorderBox() *OMathBorderBox {
	retVal, _ := this.PropGet(0x0000006c, nil)
	return NewOMathBorderBox(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Delim() *OMathDelim {
	retVal, _ := this.PropGet(0x0000006d, nil)
	return NewOMathDelim(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) EqArray() *OMathEqArray {
	retVal, _ := this.PropGet(0x0000006e, nil)
	return NewOMathEqArray(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Frac() *OMathFrac {
	retVal, _ := this.PropGet(0x0000006f, nil)
	return NewOMathFrac(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Func() *OMathFunc {
	retVal, _ := this.PropGet(0x00000070, nil)
	return NewOMathFunc(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) GroupChar() *OMathGroupChar {
	retVal, _ := this.PropGet(0x00000071, nil)
	return NewOMathGroupChar(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) LimLow() *OMathLimLow {
	retVal, _ := this.PropGet(0x00000072, nil)
	return NewOMathLimLow(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) LimUpp() *OMathLimUpp {
	retVal, _ := this.PropGet(0x00000073, nil)
	return NewOMathLimUpp(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Mat() *OMathMat {
	retVal, _ := this.PropGet(0x00000074, nil)
	return NewOMathMat(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Nary() *OMathNary {
	retVal, _ := this.PropGet(0x00000075, nil)
	return NewOMathNary(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Phantom() *OMathPhantom {
	retVal, _ := this.PropGet(0x00000076, nil)
	return NewOMathPhantom(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) ScrPre() *OMathScrPre {
	retVal, _ := this.PropGet(0x00000077, nil)
	return NewOMathScrPre(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Rad() *OMathRad {
	retVal, _ := this.PropGet(0x00000078, nil)
	return NewOMathRad(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) ScrSub() *OMathScrSub {
	retVal, _ := this.PropGet(0x00000079, nil)
	return NewOMathScrSub(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) ScrSubSup() *OMathScrSubSup {
	retVal, _ := this.PropGet(0x0000007a, nil)
	return NewOMathScrSubSup(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) ScrSup() *OMathScrSup {
	retVal, _ := this.PropGet(0x0000007b, nil)
	return NewOMathScrSup(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) OMath() *OMath {
	retVal, _ := this.PropGet(0x0000007d, nil)
	return NewOMath(retVal.IDispatch(), false, true)
}

func (this *OMathFunction) Remove() *OMathFunction {
	retVal, _ := this.Call(0x000000c9, nil)
	return NewOMathFunction(retVal.IDispatch(), false, true)
}

