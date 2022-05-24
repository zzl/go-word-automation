package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209AF-0000-0000-C000-000000000046
var IID_ReadabilityStatistic = syscall.GUID{0x000209AF, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ReadabilityStatistic struct {
	ole.OleClient
}

func NewReadabilityStatistic(pDisp *win32.IDispatch, addRef bool, scoped bool) *ReadabilityStatistic {
	 if pDisp == nil {
		return nil;
	}
	p := &ReadabilityStatistic{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ReadabilityStatisticFromVar(v ole.Variant) *ReadabilityStatistic {
	return NewReadabilityStatistic(v.IDispatch(), false, false)
}

func (this *ReadabilityStatistic) IID() *syscall.GUID {
	return &IID_ReadabilityStatistic
}

func (this *ReadabilityStatistic) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ReadabilityStatistic) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *ReadabilityStatistic) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ReadabilityStatistic) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *ReadabilityStatistic) Name() string {
	retVal, _ := this.PropGet(0x00000000, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *ReadabilityStatistic) Value() float32 {
	retVal, _ := this.PropGet(0x00000001, nil)
	return retVal.FltValVal()
}

