package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"time"
)

// 0002093D-0000-0000-C000-000000000046
var IID_Comment = syscall.GUID{0x0002093D, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Comment struct {
	ole.OleClient
}

func NewComment(pDisp *win32.IDispatch, addRef bool, scoped bool) *Comment {
	 if pDisp == nil {
		return nil;
	}
	p := &Comment{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func CommentFromVar(v ole.Variant) *Comment {
	return NewComment(v.IDispatch(), false, false)
}

func (this *Comment) IID() *syscall.GUID {
	return &IID_Comment
}

func (this *Comment) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Comment) Application() *Application {
	retVal, _ := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.IDispatch(), false, true)
}

func (this *Comment) Creator() int32 {
	retVal, _ := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Comment) Parent() *ole.DispatchClass {
	retVal, _ := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.IDispatch(), true)
}

func (this *Comment) Range() *Range {
	retVal, _ := this.PropGet(0x000003eb, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Comment) Reference() *Range {
	retVal, _ := this.PropGet(0x000003ec, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Comment) Scope() *Range {
	retVal, _ := this.PropGet(0x000003ed, nil)
	return NewRange(retVal.IDispatch(), false, true)
}

func (this *Comment) Index() int32 {
	retVal, _ := this.PropGet(0x000003ee, nil)
	return retVal.LValVal()
}

func (this *Comment) Author() string {
	retVal, _ := this.PropGet(0x000003ef, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Comment) SetAuthor(rhs string)  {
	_ = this.PropPut(0x000003ef, []interface{}{rhs})
}

func (this *Comment) Initial() string {
	retVal, _ := this.PropGet(0x000003f0, nil)
	return win32.BstrToStrAndFree(retVal.BstrValVal())
}

func (this *Comment) SetInitial(rhs string)  {
	_ = this.PropPut(0x000003f0, []interface{}{rhs})
}

func (this *Comment) ShowTip() bool {
	retVal, _ := this.PropGet(0x000003f1, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

func (this *Comment) SetShowTip(rhs bool)  {
	_ = this.PropPut(0x000003f1, []interface{}{rhs})
}

func (this *Comment) Delete()  {
	retVal, _ := this.Call(0x0000000a, nil)
	_= retVal
}

func (this *Comment) Edit()  {
	retVal, _ := this.Call(0x000003f3, nil)
	_= retVal
}

func (this *Comment) Date() time.Time {
	retVal, _ := this.PropGet(0x000003f2, nil)
	return ole.Date(retVal.DateVal()).ToGoTime()
}

func (this *Comment) IsInk() bool {
	retVal, _ := this.PropGet(0x000003f4, nil)
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

