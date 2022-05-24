package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020967-0000-0000-C000-000000000046
var IID_Bookmarks = syscall.GUID{0x00020967, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Bookmarks struct {
	ole.OleClient
}

func NewBookmarks(pDisp *win32.IDispatch, addRef bool, scoped bool) *Bookmarks {
	 if pDisp == nil {
		return nil;
	}
	p := &Bookmarks{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BookmarksFromVar(v ole.Variant) *Bookmarks {
	return NewBookmarks(v.IDispatch(), false, false)
}

func (this *Bookmarks) IID() *syscall.GUID {
	return &IID_Bookmarks
}

func (this *Bookmarks) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Bookmarks) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Bookmarks) ForEach(action func(item *Bookmark) bool) {
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
		pItem := (*Bookmark)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Bookmarks) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Bookmarks) DefaultSorting() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Bookmarks) SetDefaultSorting(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Bookmarks) ShowHidden() bool {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Bookmarks) SetShowHidden(rhs bool)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *Bookmarks) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Bookmarks) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Bookmarks) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Bookmarks) Item(index *ole.Variant) *Bookmark {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewBookmark(retVal.IDispatch(), false, true)
}

var Bookmarks_Add_OptArgs= []string{
	"Range", 
}

func (this *Bookmarks) Add(name string, optArgs ...interface{}) *Bookmark {
	optArgs = ole.ProcessOptArgs(Bookmarks_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000005, []interface{}{name}, optArgs...)
	return NewBookmark(retVal.IDispatch(), false, true)
}

func (this *Bookmarks) Exists(name string) bool {
	retVal, _ := this.Call(0x00000006, []interface{}{name})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

