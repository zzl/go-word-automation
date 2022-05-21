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
	return NewBookmarks(v.PdispValVal(), false, false)
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
	retVal := this.PropGet(-4, nil)
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
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Bookmarks) DefaultSorting() int32 {
	retVal := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Bookmarks) SetDefaultSorting(rhs int32)  {
	retVal := this.PropPut(0x00000003, []interface{}{rhs})
	_= retVal
}

func (this *Bookmarks) ShowHidden() bool {
	retVal := this.PropGet(0x00000004, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Bookmarks) SetShowHidden(rhs bool)  {
	retVal := this.PropPut(0x00000004, []interface{}{rhs})
	_= retVal
}

func (this *Bookmarks) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Bookmarks) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Bookmarks) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Bookmarks) Item(index *ole.Variant) *Bookmark {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewBookmark(retVal.PdispValVal(), false, true)
}

var Bookmarks_Add_OptArgs= []string{
	"Range", 
}

func (this *Bookmarks) Add(name string, optArgs ...interface{}) *Bookmark {
	optArgs = ole.ProcessOptArgs(Bookmarks_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000005, []interface{}{name}, optArgs...)
	return NewBookmark(retVal.PdispValVal(), false, true)
}

func (this *Bookmarks) Exists(name string) bool {
	retVal := this.Call(0x00000006, []interface{}{name})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

