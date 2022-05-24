package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209AE-0000-0000-C000-000000000046
var IID_ReadabilityStatistics = syscall.GUID{0x000209AE, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ReadabilityStatistics struct {
	ole.OleClient
}

func NewReadabilityStatistics(pDisp *win32.IDispatch, addRef bool, scoped bool) *ReadabilityStatistics {
	 if pDisp == nil {
		return nil;
	}
	p := &ReadabilityStatistics{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ReadabilityStatisticsFromVar(v ole.Variant) *ReadabilityStatistics {
	return NewReadabilityStatistics(v.IDispatch(), false, false)
}

func (this *ReadabilityStatistics) IID() *syscall.GUID {
	return &IID_ReadabilityStatistics
}

func (this *ReadabilityStatistics) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ReadabilityStatistics) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ReadabilityStatistics) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ReadabilityStatistics) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ReadabilityStatistics) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *ReadabilityStatistics) ForEach(action func(item *ReadabilityStatistic) bool) {
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
		pItem := (*ReadabilityStatistic)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *ReadabilityStatistics) Count() int32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.LValVal()
}

func (this *ReadabilityStatistics) Item(index *ole.Variant) *ReadabilityStatistic {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewReadabilityStatistic(retVal.IDispatch(), false, true)
}

