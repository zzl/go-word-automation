package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// FF06FEF2-DA89-41C0-A0A8-5CD434E210AD
var IID_ChartCharacters = syscall.GUID{0xFF06FEF2, 0xDA89, 0x41C0, 
	[8]byte{0xA0, 0xA8, 0x5C, 0xD4, 0x34, 0xE2, 0x10, 0xAD}}

type ChartCharacters struct {
	ole.OleClient
}

func NewChartCharacters(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartCharacters {
	 if pDisp == nil {
		return nil;
	}
	p := &ChartCharacters{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartCharactersFromVar(v ole.Variant) *ChartCharacters {
	return NewChartCharacters(v.IDispatch(), false, false)
}

func (this *ChartCharacters) IID() *syscall.GUID {
	return &IID_ChartCharacters
}

func (this *ChartCharacters) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartCharacters) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartCharacters) Caption() string {
	retVal, _ := this.PropGet(0x0000008b, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartCharacters) SetCaption(rhs string)  {
	_ = this.PropPut(0x0000008b, []interface{}{rhs})
}

func (this *ChartCharacters) Count() int32 {
	retVal, _ := this.PropGet(0x00000076, nil)
	return retVal.LValVal()
}

func (this *ChartCharacters) Delete() ole.Variant {
	retVal, _ := this.Call(0x00000075, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartCharacters) Font() *ChartFont {
	retVal, _ := this.PropGet(0x00000092, nil)
	return NewChartFont(retVal.IDispatch(), false, true)
}

func (this *ChartCharacters) Insert(string string) ole.Variant {
	retVal, _ := this.Call(0x000000fc, []interface{}{string})
	com.AddToScope(retVal)
	return *retVal
}

func (this *ChartCharacters) Text() string {
	retVal, _ := this.PropGet(0x0000008a, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartCharacters) SetText(rhs string)  {
	_ = this.PropPut(0x0000008a, []interface{}{rhs})
}

func (this *ChartCharacters) PhoneticCharacters() string {
	retVal, _ := this.PropGet(0x000005f2, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ChartCharacters) SetPhoneticCharacters(rhs string)  {
	_ = this.PropPut(0x000005f2, []interface{}{rhs})
}

func (this *ChartCharacters) Application() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ChartCharacters) Creator() int32 {
	retVal, _ := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

