package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 0002099B-0000-0000-C000-000000000046
var IID_SynonymInfo = syscall.GUID{0x0002099B, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type SynonymInfo struct {
	ole.OleClient
}

func NewSynonymInfo(pDisp *win32.IDispatch, addRef bool, scoped bool) *SynonymInfo {
	 if pDisp == nil {
		return nil;
	}
	p := &SynonymInfo{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func SynonymInfoFromVar(v ole.Variant) *SynonymInfo {
	return NewSynonymInfo(v.IDispatch(), false, false)
}

func (this *SynonymInfo) IID() *syscall.GUID {
	return &IID_SynonymInfo
}

func (this *SynonymInfo) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *SynonymInfo) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *SynonymInfo) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *SynonymInfo) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *SynonymInfo) Word() string {
	retVal, _ := this.PropGet(0x00000001, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *SynonymInfo) Found() bool {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *SynonymInfo) MeaningCount() int32 {
	retVal, _ := this.PropGet(0x00000003, nil)
	return retVal.LValVal()
}

func (this *SynonymInfo) MeaningList() ole.Variant {
	retVal, _ := this.PropGet(0x00000004, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *SynonymInfo) PartOfSpeechList() ole.Variant {
	retVal, _ := this.PropGet(0x00000005, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *SynonymInfo) SynonymList(meaning *ole.Variant) ole.Variant {
	retVal, _ := this.PropGet(0x00000007, []interface{}{meaning})
	com.AddToScope(retVal)
	return *retVal
}

func (this *SynonymInfo) AntonymList() ole.Variant {
	retVal, _ := this.PropGet(0x00000008, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *SynonymInfo) RelatedExpressionList() ole.Variant {
	retVal, _ := this.PropGet(0x00000009, nil)
	com.AddToScope(retVal)
	return *retVal
}

func (this *SynonymInfo) RelatedWordList() ole.Variant {
	retVal, _ := this.PropGet(0x0000000a, nil)
	com.AddToScope(retVal)
	return *retVal
}

