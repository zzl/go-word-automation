package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020973-0000-0000-C000-000000000046
var IID_TextColumns = syscall.GUID{0x00020973, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type TextColumns struct {
	ole.OleClient
}

func NewTextColumns(pDisp *win32.IDispatch, addRef bool, scoped bool) *TextColumns {
	 if pDisp == nil {
		return nil;
	}
	p := &TextColumns{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func TextColumnsFromVar(v ole.Variant) *TextColumns {
	return NewTextColumns(v.IDispatch(), false, false)
}

func (this *TextColumns) IID() *syscall.GUID {
	return &IID_TextColumns
}

func (this *TextColumns) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *TextColumns) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *TextColumns) ForEach(action func(item *TextColumn) bool) {
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
		pItem := (*TextColumn)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *TextColumns) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *TextColumns) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *TextColumns) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *TextColumns) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *TextColumns) EvenlySpaced() int32 {
	retVal, _ := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *TextColumns) SetEvenlySpaced(rhs int32)  {
	_ = this.PropPut(0x00000064, []interface{}{rhs})
}

func (this *TextColumns) LineBetween() int32 {
	retVal, _ := this.PropGet(0x00000065, nil)
	return retVal.LValVal()
}

func (this *TextColumns) SetLineBetween(rhs int32)  {
	_ = this.PropPut(0x00000065, []interface{}{rhs})
}

func (this *TextColumns) Width() float32 {
	retVal, _ := this.PropGet(0x00000066, nil)
	return retVal.FltValVal()
}

func (this *TextColumns) SetWidth(rhs float32)  {
	_ = this.PropPut(0x00000066, []interface{}{rhs})
}

func (this *TextColumns) Spacing() float32 {
	retVal, _ := this.PropGet(0x00000067, nil)
	return retVal.FltValVal()
}

func (this *TextColumns) SetSpacing(rhs float32)  {
	_ = this.PropPut(0x00000067, []interface{}{rhs})
}

func (this *TextColumns) Item(index int32) *TextColumn {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewTextColumn(retVal.IDispatch(), false, true)
}

var TextColumns_Add_OptArgs= []string{
	"Width", "Spacing", "EvenlySpaced", 
}

func (this *TextColumns) Add(optArgs ...interface{}) *TextColumn {
	optArgs = ole.ProcessOptArgs(TextColumns_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x000000c9, nil, optArgs...)
	return NewTextColumn(retVal.IDispatch(), false, true)
}

func (this *TextColumns) SetCount(numColumns int32)  {
	retVal, _ := this.Call(0x000000ca, []interface{}{numColumns})
	_= retVal
}

func (this *TextColumns) FlowDirection() int32 {
	retVal, _ := this.PropGet(0x00000068, nil)
	return retVal.LValVal()
}

func (this *TextColumns) SetFlowDirection(rhs int32)  {
	_ = this.PropPut(0x00000068, []interface{}{rhs})
}

