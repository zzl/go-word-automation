package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"time"
)

// 000209B4-0000-0000-C000-000000000046
var IID_Version = syscall.GUID{0x000209B4, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Version struct {
	ole.OleClient
}

func NewVersion(pDisp *win32.IDispatch, addRef bool, scoped bool) *Version {
	 if pDisp == nil {
		return nil;
	}
	p := &Version{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func VersionFromVar(v ole.Variant) *Version {
	return NewVersion(v.IDispatch(), false, false)
}

func (this *Version) IID() *syscall.GUID {
	return &IID_Version
}

func (this *Version) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Version) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Version) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Version) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Version) SavedBy() string {
	retVal, _ := this.PropGet(0x000003eb, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Version) Comment() string {
	retVal, _ := this.PropGet(0x000003ec, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Version) Date() time.Time {
	retVal, _ := this.PropGet(0x000003ed, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *Version) Index() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Version) OpenOld()  {
	retVal, _ := this.Call(0x00000065, nil)
	_= retVal
}

func (this *Version) Delete()  {
	retVal, _ := this.Call(0x00000066, nil)
	_= retVal
}

func (this *Version) Open() *Document {
	retVal, _ := this.Call(0x00000067, nil)
	return NewDocument(retVal.IDispatch(), false, true)
}

