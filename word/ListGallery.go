package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 00020994-0000-0000-C000-000000000046
var IID_ListGallery = syscall.GUID{0x00020994, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ListGallery struct {
	ole.OleClient
}

func NewListGallery(pDisp *win32.IDispatch, addRef bool, scoped bool) *ListGallery {
	 if pDisp == nil {
		return nil;
	}
	p := &ListGallery{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ListGalleryFromVar(v ole.Variant) *ListGallery {
	return NewListGallery(v.IDispatch(), false, false)
}

func (this *ListGallery) IID() *syscall.GUID {
	return &IID_ListGallery
}

func (this *ListGallery) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ListGallery) ListTemplates() *ListTemplates {
	retVal, _ := this.PropGet(0x00000001, nil)
	return NewListTemplates(retVal.IDispatch(), false, true)
}

func (this *ListGallery) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ListGallery) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ListGallery) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ListGallery) Modified(index int32) bool {
	retVal, _ := this.PropGet(0x00000065, []interface{}{index})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *ListGallery) Reset(index int32)  {
	retVal, _ := this.Call(0x00000064, []interface{}{index})
	_= retVal
}

