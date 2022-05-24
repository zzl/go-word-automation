package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 804CD967-F83B-432D-9446-C61A45CFEFF0
var IID_ContentControls = syscall.GUID{0x804CD967, 0xF83B, 0x432D, 
	[8]byte{0x94, 0x46, 0xC6, 0x1A, 0x45, 0xCF, 0xEF, 0xF0}}

type ContentControls struct {
	ole.OleClient
}

func NewContentControls(pDisp *win32.IDispatch, addRef bool, scoped bool) *ContentControls {
	 if pDisp == nil {
		return nil;
	}
	p := &ContentControls{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ContentControlsFromVar(v ole.Variant) *ContentControls {
	return NewContentControls(v.IDispatch(), false, false)
}

func (this *ContentControls) IID() *syscall.GUID {
	return &IID_ContentControls
}

func (this *ContentControls) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ContentControls) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ContentControls) ForEach(action func(item *ContentControl) bool) {
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
		pItem := (*ContentControl)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ContentControls) Application() *Application {
	retVal, _ := this.PropGet(0x00000064, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ContentControls) Creator() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *ContentControls) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x00000066, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ContentControls) Count() int32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *ContentControls) Item(index *ole.Variant) *ContentControl {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewContentControl(retVal.IDispatch(), false, true)
}

var ContentControls_Add_OptArgs= []string{
	"Type", "Range", 
}

func (this *ContentControls) Add(optArgs ...interface{}) *ContentControl {
	optArgs = ole.ProcessOptArgs(ContentControls_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000001, nil, optArgs...)
	return NewContentControl(retVal.IDispatch(), false, true)
}

