package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 7A1BCE11-5783-4C7D-BD02-F3D84AB40E7F
var IID_HiLoLines = syscall.GUID{0x7A1BCE11, 0x5783, 0x4C7D, 
	[8]byte{0xBD, 0x02, 0xF3, 0xD8, 0x4A, 0xB4, 0x0E, 0x7F}}

type HiLoLines struct {
	ole.OleClient
}

func NewHiLoLines(pDisp *win32.IDispatch, addRef bool, scoped bool) *HiLoLines {
	p := &HiLoLines{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HiLoLinesFromVar(v ole.Variant) *HiLoLines {
	return NewHiLoLines(v.PdispValVal(), false, false)
}

func (this *HiLoLines) IID() *syscall.GUID {
	return &IID_HiLoLines
}

func (this *HiLoLines) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *HiLoLines) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *HiLoLines) Name() string {
	retVal := this.PropGet(0x60020001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *HiLoLines) Select()  {
	retVal := this.Call(0x60020002, nil)
	_= retVal
}

func (this *HiLoLines) Border() *ChartBorder {
	retVal := this.PropGet(0x60020003, nil)
	return NewChartBorder(retVal.PdispValVal(), false, true)
}

func (this *HiLoLines) Delete()  {
	retVal := this.Call(0x60020004, nil)
	_= retVal
}

func (this *HiLoLines) Format() *ChartFormat {
	retVal := this.PropGet(0x60020005, nil)
	return NewChartFormat(retVal.PdispValVal(), false, true)
}

func (this *HiLoLines) Application() *ole.DispatchClass {
	retVal := this.PropGet(0x00000094, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *HiLoLines) Creator() int32 {
	retVal := this.PropGet(0x00000095, nil)
	return retVal.LValVal()
}

