package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 352840A9-AF7D-4CA4-87FC-21C68FDAB3E4
var IID_Page = syscall.GUID{0x352840A9, 0xAF7D, 0x4CA4, 
	[8]byte{0x87, 0xFC, 0x21, 0xC6, 0x8F, 0xDA, 0xB3, 0xE4}}

type Page struct {
	ole.OleClient
}

func NewPage(pDisp *win32.IDispatch, addRef bool, scoped bool) *Page {
	 if pDisp == nil {
		return nil;
	}
	p := &Page{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func PageFromVar(v ole.Variant) *Page {
	return NewPage(v.IDispatch(), false, false)
}

func (this *Page) IID() *syscall.GUID {
	return &IID_Page
}

func (this *Page) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Page) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Page) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Page) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Page) Left() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Page) Top() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *Page) Width() int32 {
	retVal, _ := this.PropGet(0x00000004, nil)
	return retVal.LValVal()
}

func (this *Page) Height() int32 {
	retVal, _ := this.PropGet(0x00000005, nil)
	return retVal.LValVal()
}

func (this *Page) Rectangles() *Rectangles {
	retVal, _ := this.PropGet(0x00000006, nil)
	return NewRectangles(retVal.IDispatch(), false, true)
}

func (this *Page) Breaks() *Breaks {
	retVal, _ := this.PropGet(0x00000007, nil)
	return NewBreaks(retVal.IDispatch(), false, true)
}

func (this *Page) EnhMetaFileBits() ole.Variant {
	retVal, _ := this.PropGet(0x00000008, nil)
	com.AddToScope(retVal)
	return *retVal
}

