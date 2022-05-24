package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020974-0000-0000-C000-000000000046
var IID_TextColumn = syscall.GUID{0x00020974, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TextColumn struct {
	ole.OleClient
}

func NewTextColumn(pDisp *win32.IDispatch, addRef bool, scoped bool) *TextColumn {
	 if pDisp == nil {
		return nil;
	}
	p := &TextColumn{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TextColumnFromVar(v ole.Variant) *TextColumn {
	return NewTextColumn(v.IDispatch(), false, false)
}

func (this *TextColumn) IID() *syscall.GUID {
	return &IID_TextColumn
}

func (this *TextColumn) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TextColumn) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TextColumn) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TextColumn) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextColumn) Width() float32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.FltValVal()
}

func (this *TextColumn) SetWidth(rhs float32)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *TextColumn) SpaceAfter() float32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.FltValVal()
}

func (this *TextColumn) SetSpaceAfter(rhs float32)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

