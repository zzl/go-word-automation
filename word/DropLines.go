package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 9F1DF642-3CCE-4D83-A770-D2634A05D278
var IID_DropLines = syscall.GUID{0x9F1DF642, 0x3CCE, 0x4D83, 
	[8]byte{0xA7, 0x70, 0xD2, 0x63, 0x4A, 0x05, 0xD2, 0x78}}

type DropLines struct {
	ole.OleClient
}

func NewDropLines(pDisp *win32.IDispatch, addRef bool, scoped bool) *DropLines {
	p := &DropLines{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DropLinesFromVar(v ole.Variant) *DropLines {
	return NewDropLines(v.PdispValVal(), false, false)
}

func (this *DropLines) IID() *syscall.GUID {
	return &IID_DropLines
}

func (this *DropLines) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *DropLines) Name() string {
	retVal := this.PropGet(0x60020000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *DropLines) Select()  {
	retVal := this.Call(0x60020001, nil)
	_= retVal
}

func (this *DropLines) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x60020002, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DropLines) Border() *ChartBorder {
	retVal := this.PropGet(0x60020003, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *DropLines) Delete()  {
	retVal := this.Call(0x60020004, nil)
	_= retVal
}

func (this *DropLines) Format() *ChartFormat {
	retVal := this.PropGet(0x60020005, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *DropLines) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *DropLines) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

