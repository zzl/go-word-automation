package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209EE-0000-0000-C000-000000000046
var IID_SmartTags = syscall.GUID{0x000209EE, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SmartTags struct {
	ole.OleClient
}

func NewSmartTags(pDisp *win32.IDispatch, addRef bool, scoped bool) *SmartTags {
	 if pDisp == nil {
		return nil;
	}
	p := &SmartTags{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SmartTagsFromVar(v ole.Variant) *SmartTags {
	return NewSmartTags(v.IDispatch(), false, false)
}

func (this *SmartTags) IID() *syscall.GUID {
	return &IID_SmartTags
}

func (this *SmartTags) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SmartTags) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *SmartTags) ForEach(action func(item *SmartTag) bool) {
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
		pItem := (*SmartTag)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *SmartTags) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *SmartTags) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SmartTags) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SmartTags) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SmartTags) Item(index *ole.Variant) *SmartTag {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewSmartTag(retVal.IDispatch(), false, true)
}

var SmartTags_Add_OptArgs= []string{
	"Range", "Properties", 
}

func (this *SmartTags) Add(name string, optArgs ...interface{}) *SmartTag {
	optArgs = ole.ProcessOptArgs(SmartTags_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000005, []interface{}{name}, optArgs...)
	return NewSmartTag(retVal.IDispatch(), false, true)
}

func (this *SmartTags) SmartTagsByType(name string) *SmartTags {
	retVal, _ := this.Call(0x000003eb, []interface{}{name})
	return NewSmartTags(retVal.IDispatch(), false, true)
}

