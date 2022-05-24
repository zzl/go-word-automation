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
	 if pDisp == nil {
		return nil;
	}
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
	return NewUndoRecord(v.IDispatch(), false, false)
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
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *UndoRecord) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *UndoRecord) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

var UndoRecord_StartCustomRecord_OptArgs= []string{
	"Name", 
}

func (this *UndoRecord) StartCustomRecord(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(UndoRecord_StartCustomRecord_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000001, nil, optArgs...)
	_= retVal
}

func (this *UndoRecord) EndCustomRecord()  {
	retVal, _ := this.Call(0x00000002, nil)
	_= retVal
}

func (this *UndoRecord) IsRecordingCustomRecord() bool {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *UndoRecord) CustomRecordName() string {
	retVal, _ := this.PropGet(0x00000004, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *UndoRecord) CustomRecordLevel() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

