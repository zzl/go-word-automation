package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 0002096C-0000-0000-C000-000000000046
var IID_Documents = syscall.GUID{0x0002096C, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type Documents struct {
	ole.OleClient
}

func NewDocuments(pDisp *win32.IDispatch, addRef bool, scoped bool) *Documents {
	p := &Documents{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func DocumentsFromVar(v ole.Variant) *Documents {
	return NewDocuments(v.PdispValVal(), false, false)
}

func (this *Documents) IID() *syscall.GUID {
	return &IID_Documents
}

func (this *Documents) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *Documents) NewEnum_() *com.UnknownClass {
	retVal := this.PropGet(-4, nil)
	return com.NewUnknownClass(retVal.PunkValVal(), true)
}

func (this *Documents) ForEach(action func(item *Document) bool) {
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
		pItem := (*Document)(v.ToPointer())
		ret := action(pItem)
		v.Clear()
		if !ret {
			break
		}
	}
}

func (this *Documents) Count() int32 {
	retVal := this.PropGet(0x00000002, nil)
	return retVal.LValVal()
}

func (this *Documents) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *Documents) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *Documents) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x000003ea, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *Documents) Item(index *ole.Variant) *Document {
	retVal := this.Call(0x00000000, []interface{}{index})
	return NewDocument(retVal.PdispValVal(), false, true)
}

var Documents_Close_OptArgs= []string{
	"SaveChanges", "OriginalFormat", "RouteDocument", 
}

func (this *Documents) Close(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Documents_Close_OptArgs, optArgs)
	retVal := this.Call(0x00000451, nil, optArgs...)
	_= retVal
}

var Documents_AddOld_OptArgs= []string{
	"Template", "NewTemplate", 
}

func (this *Documents) AddOld(optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(Documents_AddOld_OptArgs, optArgs)
	retVal := this.Call(0x0000000b, nil, optArgs...)
	return NewDocument(retVal.PdispValVal(), false, true)
}

var Documents_OpenOld_OptArgs= []string{
	"ConfirmConversions", "ReadOnly", "AddToRecentFiles", "PasswordDocument", 
	"PasswordTemplate", "Revert", "WritePasswordDocument", "WritePasswordTemplate", "Format", 
}

func (this *Documents) OpenOld(fileName *ole.Variant, optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(Documents_OpenOld_OptArgs, optArgs)
	retVal := this.Call(0x0000000c, []interface{}{fileName}, optArgs...)
	return NewDocument(retVal.PdispValVal(), false, true)
}

var Documents_Save_OptArgs= []string{
	"NoPrompt", "OriginalFormat", 
}

func (this *Documents) Save(optArgs ...interface{})  {
	optArgs = ole.ProcessOptArgs(Documents_Save_OptArgs, optArgs)
	retVal := this.Call(0x0000000d, nil, optArgs...)
	_= retVal
}

var Documents_Add_OptArgs= []string{
	"Template", "NewTemplate", "DocumentType", "Visible", 
}

func (this *Documents) Add(optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(Documents_Add_OptArgs, optArgs)
	retVal := this.Call(0x0000000e, nil, optArgs...)
	return NewDocument(retVal.PdispValVal(), false, true)
}

var Documents_Open2000_OptArgs= []string{
	"ConfirmConversions", "ReadOnly", "AddToRecentFiles", "PasswordDocument", 
	"PasswordTemplate", "Revert", "WritePasswordDocument", "WritePasswordTemplate", 
	"Format", "Encoding", "Visible", 
}

func (this *Documents) Open2000(fileName *ole.Variant, optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(Documents_Open2000_OptArgs, optArgs)
	retVal := this.Call(0x0000000f, []interface{}{fileName}, optArgs...)
	return NewDocument(retVal.PdispValVal(), false, true)
}

func (this *Documents) CheckOut(fileName string)  {
	retVal := this.Call(0x00000010, []interface{}{fileName})
	_= retVal
}

func (this *Documents) CanCheckOut(fileName string) bool {
	retVal := this.Call(0x00000011, []interface{}{fileName})
	return retVal.BoolValVal() != win32.VARIANT_FALSE
}

var Documents_Open2002_OptArgs= []string{
	"ConfirmConversions", "ReadOnly", "AddToRecentFiles", "PasswordDocument", 
	"PasswordTemplate", "Revert", "WritePasswordDocument", "WritePasswordTemplate", 
	"Format", "Encoding", "Visible", "OpenAndRepair", 
	"DocumentDirection", "NoEncodingDialog", 
}

func (this *Documents) Open2002(fileName *ole.Variant, optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(Documents_Open2002_OptArgs, optArgs)
	retVal := this.Call(0x00000012, []interface{}{fileName}, optArgs...)
	return NewDocument(retVal.PdispValVal(), false, true)
}

var Documents_Open_OptArgs= []string{
	"ConfirmConversions", "ReadOnly", "AddToRecentFiles", "PasswordDocument", 
	"PasswordTemplate", "Revert", "WritePasswordDocument", "WritePasswordTemplate", 
	"Format", "Encoding", "Visible", "OpenAndRepair", 
	"DocumentDirection", "NoEncodingDialog", "XMLTransform", 
}

func (this *Documents) Open(fileName *ole.Variant, optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(Documents_Open_OptArgs, optArgs)
	retVal := this.Call(0x00000013, []interface{}{fileName}, optArgs...)
	return NewDocument(retVal.PdispValVal(), false, true)
}

var Documents_OpenNoRepairDialog_OptArgs= []string{
	"ConfirmConversions", "ReadOnly", "AddToRecentFiles", "PasswordDocument", 
	"PasswordTemplate", "Revert", "WritePasswordDocument", "WritePasswordTemplate", 
	"Format", "Encoding", "Visible", "OpenAndRepair", 
	"DocumentDirection", "NoEncodingDialog", "XMLTransform", 
}

func (this *Documents) OpenNoRepairDialog(fileName *ole.Variant, optArgs ...interface{}) *Document {
	optArgs = ole.ProcessOptArgs(Documents_OpenNoRepairDialog_OptArgs, optArgs)
	retVal := this.Call(0x00000014, []interface{}{fileName}, optArgs...)
	return NewDocument(retVal.PdispValVal(), false, true)
}

func (this *Documents) AddBlogDocument(providerID string, postURL string, blogName string, postID string) *Document {
	retVal := this.Call(0x00000015, []interface{}{providerID, postURL, blogName, postID})
	return NewDocument(retVal.PdispValVal(), false, true)
}

