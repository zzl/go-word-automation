package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020964-0000-0000-C000-000000000046
var IID_RecentFile = syscall.GUID{0x00020964, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type RecentFile struct {
	ole.OleClient
}

func NewRecentFile(pDisp *win32.IDispatch, addRef bool, scoped bool) *RecentFile {
	 if pDisp == nil {
		return nil;
	}
	p := &RecentFile{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RecentFileFromVar(v ole.Variant) *RecentFile {
	return NewRecentFile(v.IDispatch(), false, false)
}

func (this *RecentFile) IID() *syscall.GUID {
	return &IID_RecentFile
}

func (this *RecentFile) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *RecentFile) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *RecentFile) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *RecentFile) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *RecentFile) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *RecentFile) Index() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *RecentFile) ReadOnly() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *RecentFile) SetReadOnly(rhs bool)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *RecentFile) Path() string {
	retVal, _ := this.PropGet(0x00000003, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *RecentFile) Open() *Document {
	retVal, _ := this.Call(0x00000004, nil)
	return NewDocument(retVal.IDispatch(), false, true)
}

func (this *RecentFile) Delete()  {
	retVal, _ := this.Call(0x00000005, nil)
	_= retVal
}

