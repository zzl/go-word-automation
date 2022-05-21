package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 5D7F6C15-36CE-44CC-9692-5A1F8B8C906D
var IID_SeriesLines = syscall.GUID{0x5D7F6C15, 0x36CE, 0x44CC, 
	[8]byte{0x96, 0x92, 0x5A, 0x1F, 0x8B, 0x8C, 0x90, 0x6D}}

type SeriesLines struct {
	ole.OleClient
}

func NewSeriesLines(pDisp *win32.IDispatch, addRef bool, scoped bool) *SeriesLines {
	p := &SeriesLines{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SeriesLinesFromVar(v ole.Variant) *SeriesLines {
	return NewSeriesLines(v.PdispValVal(), false, false)
}

func (this *SeriesLines) IID() *syscall.GUID {
	return &IID_SeriesLines
}

func (this *SeriesLines) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SeriesLines) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SeriesLines) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *SeriesLines) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SeriesLines) Border() *ChartBorder {
	retVal := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *SeriesLines) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *SeriesLines) Format() *ChartFormat {
	retVal := this.PropGet(0x60020005, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *SeriesLines) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *SeriesLines) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

