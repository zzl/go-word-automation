package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 36162C62-B59A-4278-AF3D-F2AC1EB999D9
var IID_LeaderLines = syscall.GUID{0x36162C62, 0xB59A, 0x4278, 
	[8]byte{0xAF, 0x3D, 0xF2, 0xAC, 0x1E, 0xB9, 0x99, 0xD9}}

type LeaderLines struct {
	ole.OleClient
}

func NewLeaderLines(pDisp *win32.IDispatch, addRef bool, scoped bool) *LeaderLines {
	p := &LeaderLines{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func LeaderLinesFromVar(v ole.Variant) *LeaderLines {
	return NewLeaderLines(v.PdispValVal(), false, false)
}

func (this *LeaderLines) IID() *syscall.GUID {
	return &IID_LeaderLines
}

func (this *LeaderLines) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *LeaderLines) Select()  {
	retVal := this.Call(0x000000eb, nil)
	_= retVal
}

func (this *LeaderLines) Border() *ChartBorder {
	retVal := this.PropGet(0x00000080, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *LeaderLines) Delete()  {
	retVal := this.Call(0x00000075, nil)
	_= retVal
}

func (this *LeaderLines) Format() *ChartFormat {
	retVal := this.PropGet(0x60020003, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *LeaderLines) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *LeaderLines) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

func (this *LeaderLines) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000096, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

