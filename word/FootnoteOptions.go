package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// BEA85A24-D7DA-4F3D-B58C-ED90FB01D615
var IID_FootnoteOptions = syscall.GUID{0xBEA85A24, 0xD7DA, 0x4F3D, 
	[8]byte{0xB5, 0x8C, 0xED, 0x90, 0xFB, 0x01, 0xD6, 0x15}}

type FootnoteOptions struct {
	ole.OleClient
}

func NewFootnoteOptions(pDisp *win32.IDispatch, addRef bool, scoped bool) *FootnoteOptions {
	 if pDisp == nil {
		return nil;
	}
	p := &FootnoteOptions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FootnoteOptionsFromVar(v ole.Variant) *FootnoteOptions {
	return NewFootnoteOptions(v.IDispatch(), false, false)
}

func (this *FootnoteOptions) IID() *syscall.GUID {
	return &IID_FootnoteOptions
}

func (this *FootnoteOptions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FootnoteOptions) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *FootnoteOptions) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *FootnoteOptions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *FootnoteOptions) Location() int32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *FootnoteOptions) SetLocation(rhs int32)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *FootnoteOptions) NumberStyle() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *FootnoteOptions) SetNumberStyle(rhs int32)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *FootnoteOptions) StartingNumber() int32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *FootnoteOptions) SetStartingNumber(rhs int32)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *FootnoteOptions) NumberingRule() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *FootnoteOptions) SetNumberingRule(rhs int32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

