package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020963-0000-0000-C000-000000000046
var IID_RecentFiles = syscall.GUID{0x00020963, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type RecentFiles struct {
	ole.OleClient
}

func NewRecentFiles(pDisp *win32.IDispatch, addRef bool, scoped bool) *RecentFiles {
	 if pDisp == nil {
		return nil;
	}
	p := &RecentFiles{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func RecentFilesFromVar(v ole.Variant) *RecentFiles {
	return NewRecentFiles(v.IDispatch(), false, false)
}

func (this *RecentFiles) IID() *syscall.GUID {
	return &IID_RecentFiles
}

func (this *RecentFiles) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *RecentFiles) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *RecentFiles) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *RecentFiles) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *RecentFiles) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *RecentFiles) ForEach(action func(item *RecentFile) bool) {
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
		pItem := (*RecentFile)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *RecentFiles) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *RecentFiles) Maximum() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *RecentFiles) SetMaximum(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *RecentFiles) Item(index int32) *RecentFile {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewRecentFile(retVal.IDispatch(), false, true)
}

var RecentFiles_Add_OptArgs= []string{
	"ReadOnly", 
}

func (this *RecentFiles) Add(document *ole.Variant, optArgs ...interface{}) *RecentFile {
	optArgs = ole.ProcessOptArgs(RecentFiles_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000003, []interface{}{document}, optArgs...)
	return NewRecentFile(retVal.IDispatch(), false, true)
}

