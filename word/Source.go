package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 4A6AE865-199D-4EA3-9F6B-125BD9C40EDF
var IID_Source = syscall.GUID{0x4A6AE865, 0x199D, 0x4EA3, 
	[8]byte{0x9F, 0x6B, 0x12, 0x5B, 0xD9, 0xC4, 0x0E, 0xDF}}

type Source struct {
	ole.OleClient
}

func NewSource(pDisp *win32.IDispatch, addRef bool, scoped bool) *Source {
	 if pDisp == nil {
		return nil;
	}
	p := &Source{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SourceFromVar(v ole.Variant) *Source {
	return NewSource(v.IDispatch(), false, false)
}

func (this *Source) IID() *syscall.GUID {
	return &IID_Source
}

func (this *Source) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Source) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Source) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *Source) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Source) Tag() string {
	retVal, _ := this.PropGet(0x00000067, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Source) Field(name string) string {
	retVal, _ := this.PropGet(0x00000068, []interface{}{name})
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Source) SetField(name string, rhs string)  {
	_ = this.PropPut(0x00000068, []interface{}{name, rhs})
}

func (this *Source) XML() string {
	retVal, _ := this.PropGet(0x00000069, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Source) Cited() bool {
	retVal, _ := this.PropGet(0x0000006b, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Source) Delete()  {
	retVal, _ := this.Call(0x0000006a, nil)
	_= retVal
}

