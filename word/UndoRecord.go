package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// E598E358-2852-42D4-8775-160BD91B7244
var IID_UndoRecord = syscall.GUID{0xE598E358, 0x2852, 0x42D4, 
	[8]byte{0x87, 0x75, 0x16, 0x0B, 0xD9, 0x1B, 0x72, 0x44}}

type UndoRecord struct {
	ole.OleClient
}

func NewUndoRecord(pDisp *win32.IDispatch, addRef bool, scoped bool) *UndoRecord {
	p := &UndoRecord{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func UndoRecordFromVar(v ole.Variant) *UndoRecord {
	return NewUndoRecord(v.PdispValVal(), false, false)
}

func (this *UndoRecord) IID() *syscall.GUID {
	return &IID_UndoRecord
}

func (this *UndoRecord) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *UndoRecord) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *UndoRecord) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *UndoRecord) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *UndoRecord) StartCustomRecord(name string)  {
	retVal := this.Call(0x00000001, []interface{}{name})
	_= retVal
}

func (this *UndoRecord) EndCustomRecord()  {
	retVal := this.Call(0x00000002, nil)
	_= retVal
}

func (this *UndoRecord) IsRecordingCustomRecord() bool {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *UndoRecord) CustomRecordName() string {
	retVal := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *UndoRecord) CustomRecordLevel() int32 {
	retVal := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

