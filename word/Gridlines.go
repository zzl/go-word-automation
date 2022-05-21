package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// FC9090AF-0DDB-4EC1-86E8-8751F2199F2C
var IID_Gridlines = syscall.GUID{0xFC9090AF, 0x0DDB, 0x4EC1, 
	[8]byte{0x86, 0xE8, 0x87, 0x51, 0xF2, 0x19, 0x9F, 0x2C}}

type Gridlines struct {
	ole.OleClient
}

func NewGridlines(pDisp *win32.IDispatch, addRef bool, scoped bool) *Gridlines {
	p := &Gridlines{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func GridlinesFromVar(v ole.Variant) *Gridlines {
	return NewGridlines(v.PdispValVal(), false, false)
}

func (this *Gridlines) IID() *syscall.GUID {
	return &IID_Gridlines
}

func (this *Gridlines) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Gridlines) Name() string {
	retVal := this.PropGet(0x0000006e, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Gridlines) Select() ole.Variant {
	retVal := this.Call(0x000000eb, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Gridlines) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Gridlines) Border() *ChartBorder {
	retVal := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *Gridlines) Delete() ole.Variant {
	retVal := this.Call(0x00000075, nil)
	com.CurrentScope.AddVarIfNeeded((*win32.VARIANT)(retVal))
	return *retVal
}

func (this *Gridlines) Format() *ChartFormat {
	retVal := this.PropGet(0x60020005, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *Gridlines) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Gridlines) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

