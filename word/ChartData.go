package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 4A304B59-31FF-42DD-B436-7FC9C5DB7559
var IID_ChartData = syscall.GUID{0x4A304B59, 0x31FF, 0x42DD, 
	[8]byte{0xB4, 0x36, 0x7F, 0xC9, 0xC5, 0xDB, 0x75, 0x59}}

type ChartData struct {
	ole.OleClient
}

func NewChartData(pDisp *win32.IDispatch, addRef bool, scoped bool) *ChartData {
	p := &ChartData{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ChartDataFromVar(v ole.Variant) *ChartData {
	return NewChartData(v.PdispValVal(), false, false)
}

func (this *ChartData) IID() *syscall.GUID {
	return &IID_ChartData
}

func (this *ChartData) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ChartData) Workbook() *ole.DispatchClass {
	retVal := this.PropGet(0x60020000, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ChartData) Activate()  {
	retVal := this.Call(0x60020001, nil)
	_= retVal
}

func (this *ChartData) IsLinked() bool {
	retVal := this.PropGet(0x60020002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ChartData) BreakLink()  {
	retVal := this.Call(0x60020003, nil)
	_= retVal
}

