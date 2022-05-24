package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020940-0000-0000-C000-000000000046
var IID_Comments = syscall.GUID{0x00020940, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Comments struct {
	ole.OleClient
}

func NewComments(pDisp *win32.IDispatch, addRef bool, scoped bool) *Comments {
	 if pDisp == nil {
		return nil;
	}
	p := &Comments{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CommentsFromVar(v ole.Variant) *Comments {
	return NewComments(v.IDispatch(), false, false)
}

func (this *Comments) IID() *syscall.GUID {
	return &IID_Comments
}

func (this *Comments) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Comments) NewEnum_() *com.UnknownClass {
	retVal, _ := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Comments) ForEach(action func(item *Comment) bool) {
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
		pItem := (*Comment)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Comments) Count() int32 {
	retVal, _ := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Comments) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Comments) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Comments) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Comments) ShowBy() string {
	retVal, _ := this.PropGet(0x000003eb, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Comments) SetShowBy(rhs string)  {
	_ = this.PropPut(0x000003eb, []interface{}{rhs})
}

func (this *Comments) Item(index int32) *Comment {
	retVal, _ := this.Call(0x00000000, []interface{}{index})
	return NewComment(retVal.IDispatch(), false, true)
}

var Comments_Add_OptArgs= []string{
	"Text", 
}

func (this *Comments) Add(range_ *Range, optArgs ...interface{}) *Comment {
	optArgs = ole.ProcessOptArgs(Comments_Add_OptArgs, optArgs)
	retVal, _ := this.Call(0x00000004, []interface{}{range_}, optArgs...)
	return NewComment(retVal.IDispatch(), false, true)
}

