package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002099A-0000-0000-C000-000000000046
var IID_FileConverters = syscall.GUID{0x0002099A, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type FileConverters struct {
	ole.OleClient
}

func NewFileConverters(pDisp *win32.IDispatch, addRef bool, scoped bool) *FileConverters {
	 if pDisp == nil {
		return nil;
	}
	p := &FileConverters{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func FileConvertersFromVar(v ole.Variant) *FileConverters {
	return NewFileConverters(v.IDispatch(), false, false)
}

func (this *FileConverters) IID() *syscall.GUID {
	return &IID_FileConverters
}

func (this *FileConverters) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *FileConverters) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *FileConverters) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *FileConverters) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *FileConverters) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *FileConverters) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *FileConverters) ForEach(action func(item *FileConverter) bool) {
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
		pItem := (*FileConverter)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *FileConverters) ConvertMacWordChevrons() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *FileConverters) SetConvertMacWordChevrons(rhs int32)  {
	_ = this.PropPut(0x00000002, []interface{}{rhs})
}

func (this *FileConverters) Item(index *ole.Variant) *FileConverter {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewFileConverter(retVal.IDispatch(), false, true)
}

