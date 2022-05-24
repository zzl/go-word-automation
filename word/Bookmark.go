package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020968-0000-0000-C000-000000000046
var IID_Bookmark = syscall.GUID{0x00020968, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Bookmark struct {
	ole.OleClient
}

func NewBookmark(pDisp *win32.IDispatch, addRef bool, scoped bool) *Bookmark {
	 if pDisp == nil {
		return nil;
	}
	p := &Bookmark{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func BookmarkFromVar(v ole.Variant) *Bookmark {
	return NewBookmark(v.IDispatch(), false, false)
}

func (this *Bookmark) IID() *syscall.GUID {
	return &IID_Bookmark
}

func (this *Bookmark) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Bookmark) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Bookmark) Range() *Range {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Bookmark) Empty() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Bookmark) Start() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Bookmark) SetStart(rhs int32)  {
	_ = this.PropPut(0x00000003, []interface{}{rhs})
}

func (this *Bookmark) End() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Bookmark) SetEnd(rhs int32)  {
	_ = this.PropPut(0x00000004, []interface{}{rhs})
}

func (this *Bookmark) Column() bool {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Bookmark) StoryType() int32 {
	retVal, _ := this.PropGet(0x00000006, nil)
	return retVal.LValVal()
}

func (this *Bookmark) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Bookmark) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Bookmark) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Bookmark) Select()  {
	retVal, _ := this.Call(0x0000ffff, nil)
	_= retVal
}

func (this *Bookmark) Delete()  {
	retVal, _ := this.Call(0x0000000b, nil)
	_= retVal
}

func (this *Bookmark) Copy(name string) *Bookmark {
	retVal, _ := this.Call(0x0000000c, []interface{}{name})
	return NewBookmark(retVal.IDispatch(), false, true)
}

