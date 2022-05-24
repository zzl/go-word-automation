package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// FD0A74E8-C719-49F6-BA1B-F6D9839D1AB9
var IID_ProtectedViewWindows = syscall.GUID{0xFD0A74E8, 0xC719, 0x49F6, 
	[8]byte{0xBA, 0x1B, 0xF6, 0xD9, 0x83, 0x9D, 0x1A, 0xB9}}

type ProtectedViewWindows struct {
	ole.OleClient
}

func NewProtectedViewWindows(pDisp *win32.IDispatch, addRef bool, scoped bool) *ProtectedViewWindows {
	 if pDisp == nil {
		return nil;
	}
	p := &ProtectedViewWindows{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ProtectedViewWindowsFromVar(v ole.Variant) *ProtectedViewWindows {
	return NewProtectedViewWindows(v.IDispatch(), false, false)
}

func (this *ProtectedViewWindows) IID() *syscall.GUID {
	return &IID_ProtectedViewWindows
}

func (this *ProtectedViewWindows) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ProtectedViewWindows) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ProtectedViewWindows) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindows) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ProtectedViewWindows) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ProtectedViewWindows) ForEach(action func(item *ProtectedViewWindow) bool) {
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
		pItem := (*ProtectedViewWindow)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ProtectedViewWindows) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *ProtectedViewWindows) Item(index *ole.Variant) *ProtectedViewWindow {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewProtectedViewWindow(retVal.IDispatch(), false, true)
}

var ProtectedViewWindows_Open_OptArgs= []string{
	"AddToRecentFiles", "PasswordDocument", "Visible", "OpenAndRepair", 
}

func (this *ProtectedViewWindows) Open(fileName *ole.Variant, optArgs ...interface{}) *ProtectedViewWindow {
	optArgs = ole.ProcessOptArgs(ProtectedViewWindows_Open_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000002, []interface{}{fileName}, optArgs...)
	return NewProtectedViewWindow(retVal.IDispatch(), false, true)
}

