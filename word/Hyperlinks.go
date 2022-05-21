package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002099C-0000-0000-C000-000000000046
var IID_Hyperlinks = syscall.GUID{0x0002099C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Hyperlinks struct {
	ole.OleClient
}

func NewHyperlinks(pDisp *win32.IDispatch, addRef bool, scoped bool) *Hyperlinks {
	p := &Hyperlinks{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func HyperlinksFromVar(v ole.Variant) *Hyperlinks {
	return NewHyperlinks(v.PdispValVal(), false, false)
}

func (this *Hyperlinks) IID() *syscall.GUID {
	return &IID_Hyperlinks
}

func (this *Hyperlinks) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Hyperlinks) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Hyperlinks) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Hyperlinks) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Hyperlinks) Count() int32 {
	retVal := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *Hyperlinks) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Hyperlinks) ForEach(action func(item *Hyperlink) bool) {
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
		pItem := (*Hyperlink)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Hyperlinks) Item(index *ole.Variant) *Hyperlink {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewHyperlink(retVal.PdispValVal(), false, true)
}

var Hyperlinks_Add__OptArgs= []string{
	"Address", "SubAddress", 
}

func (this *Hyperlinks) Add_(anchor *ole.DispatchClass, optArgs ...interface{}) *Hyperlink {
	optArgs = ole.ProcessOptArgs(Hyperlinks_Add__OptArgs, optArgs)
	retVal := this.Call(0x00000064, []interface{}{anchor}, optArgs...)
	return NewHyperlink(retVal.PdispValVal(), false, true)
}

var Hyperlinks_Add_OptArgs= []string{
	"Address", "SubAddress", "ScreenTip", "TextToDisplay", "Target", 
}

func (this *Hyperlinks) Add(anchor *ole.DispatchClass, optArgs ...interface{}) *Hyperlink {
	optArgs = ole.ProcessOptArgs(Hyperlinks_Add_OptArgs, optArgs)
	retVal := this.Call(0x00000065, []interface{}{anchor}, optArgs...)
	return NewHyperlink(retVal.PdispValVal(), false, true)
}

