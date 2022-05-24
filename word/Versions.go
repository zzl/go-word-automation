package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209B3-0000-0000-C000-000000000046
var IID_Versions = syscall.GUID{0x000209B3, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Versions struct {
	ole.OleClient
}

func NewVersions(pDisp *win32.IDispatch, addRef bool, scoped bool) *Versions {
	 if pDisp == nil {
		return nil;
	}
	p := &Versions{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func VersionsFromVar(v ole.Variant) *Versions {
	return NewVersions(v.IDispatch(), false, false)
}

func (this *Versions) IID() *syscall.GUID {
	return &IID_Versions
}

func (this *Versions) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Versions) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Versions) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Versions) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Versions) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Versions) ForEach(action func(item *Version) bool) {
	pEnum := this.NewEnum_()
	var pEnumVar *win32.IEnumVARIANT
	pEnum.QueryInterface(&win32.IID_IEnumVARIANT, unsafe.Pointer(&pEnumVar))
	defer pEnumVar.Release();
	for {
		var c uint32
		var v ole.Variant
		pEnumVar.Next(1, (*win32.VARIANT)(&v), &c)
		if c == 0 {
			break
		}
		pItem := (*Version)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Versions) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Versions) AutoVersion() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Versions) SetAutoVersion(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Versions) Item(index int32) *Version {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewVersion(retVal.IDispatch(), false, true)
}

var Versions_Save_OptArgs= []string{
	"Comment", 
}

func (this *Versions) Save(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Versions_Save_OptArgs, optArgs)
	retVal, _ := this.Call(0x0000000b, nil, optArgs...)
	_= retVal
}

