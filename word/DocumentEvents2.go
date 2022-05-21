package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 00020A02-0000-0000-C000-000000000046
var IID_DocumentEvents2 = syscall.GUID{0x00020A02, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DocumentEvents2DispInterface interface {
	New() 
	Open() 
	Close() 
	Sync(syncEventType int32) 
	XMLAfterInsert(newXMLNode *XMLNode, inUndoRedo bool) 
	XMLBeforeDelete(deletedRange *Range, oldXMLNode *XMLNode, inUndoRedo bool) 
	ContentControlAfterAdd(newContentControl *ContentControl, inUndoRedo bool) 
	ContentControlBeforeDelete(oldContentControl *ContentControl, inUndoRedo bool) 
	ContentControlOnExit(contentControl *ContentControl, cancel *win32.VARIANT_BOOL) 
	ContentControlOnEnter(contentControl *ContentControl) 
	ContentControlBeforeStoreUpdate(contentControl *ContentControl, content *win32.BSTR) 
	ContentControlBeforeContentUpdate(contentControl *ContentControl, content *win32.BSTR) 
	BuildingBlockInsert(range_ *Range, name string, category string, blockType string, template string) 
}

type DocumentEvents2Handlers struct {
	New func() 
	Open func() 
	Close func() 
	Sync func(syncEventType int32) 
	XMLAfterInsert func(newXMLNode *XMLNode, inUndoRedo bool) 
	XMLBeforeDelete func(deletedRange *Range, oldXMLNode *XMLNode, inUndoRedo bool) 
	ContentControlAfterAdd func(newContentControl *ContentControl, inUndoRedo bool) 
	ContentControlBeforeDelete func(oldContentControl *ContentControl, inUndoRedo bool) 
	ContentControlOnExit func(contentControl *ContentControl, cancel *win32.VARIANT_BOOL) 
	ContentControlOnEnter func(contentControl *ContentControl) 
	ContentControlBeforeStoreUpdate func(contentControl *ContentControl, content *win32.BSTR) 
	ContentControlBeforeContentUpdate func(contentControl *ContentControl, content *win32.BSTR) 
	BuildingBlockInsert func(range_ *Range, name string, category string, blockType string, template string) 
}

type DocumentEvents2DispImpl struct {
	Handlers DocumentEvents2Handlers
}

func (this *DocumentEvents2DispImpl) New() {
	if this.Handlers.New != nil {
		this.Handlers.New()
	}
}

func (this *DocumentEvents2DispImpl) Open() {
	if this.Handlers.Open != nil {
		this.Handlers.Open()
	}
}

func (this *DocumentEvents2DispImpl) Close() {
	if this.Handlers.Close != nil {
		this.Handlers.Close()
	}
}

func (this *DocumentEvents2DispImpl) Sync(syncEventType int32) {
	if this.Handlers.Sync != nil {
		this.Handlers.Sync(syncEventType)
	}
}

func (this *DocumentEvents2DispImpl) XMLAfterInsert(newXMLNode *XMLNode, inUndoRedo bool) {
	if this.Handlers.XMLAfterInsert != nil {
		this.Handlers.XMLAfterInsert(newXMLNode, inUndoRedo)
	}
}

func (this *DocumentEvents2DispImpl) XMLBeforeDelete(deletedRange *Range, oldXMLNode *XMLNode, inUndoRedo bool) {
	if this.Handlers.XMLBeforeDelete != nil {
		this.Handlers.XMLBeforeDelete(deletedRange, oldXMLNode, inUndoRedo)
	}
}

func (this *DocumentEvents2DispImpl) ContentControlAfterAdd(newContentControl *ContentControl, inUndoRedo bool) {
	if this.Handlers.ContentControlAfterAdd != nil {
		this.Handlers.ContentControlAfterAdd(newContentControl, inUndoRedo)
	}
}

func (this *DocumentEvents2DispImpl) ContentControlBeforeDelete(oldContentControl *ContentControl, inUndoRedo bool) {
	if this.Handlers.ContentControlBeforeDelete != nil {
		this.Handlers.ContentControlBeforeDelete(oldContentControl, inUndoRedo)
	}
}

func (this *DocumentEvents2DispImpl) ContentControlOnExit(contentControl *ContentControl, cancel *win32.VARIANT_BOOL) {
	if this.Handlers.ContentControlOnExit != nil {
		this.Handlers.ContentControlOnExit(contentControl, cancel)
	}
}

func (this *DocumentEvents2DispImpl) ContentControlOnEnter(contentControl *ContentControl) {
	if this.Handlers.ContentControlOnEnter != nil {
		this.Handlers.ContentControlOnEnter(contentControl)
	}
}

func (this *DocumentEvents2DispImpl) ContentControlBeforeStoreUpdate(contentControl *ContentControl, content *win32.BSTR) {
	if this.Handlers.ContentControlBeforeStoreUpdate != nil {
		this.Handlers.ContentControlBeforeStoreUpdate(contentControl, content)
	}
}

func (this *DocumentEvents2DispImpl) ContentControlBeforeContentUpdate(contentControl *ContentControl, content *win32.BSTR) {
	if this.Handlers.ContentControlBeforeContentUpdate != nil {
		this.Handlers.ContentControlBeforeContentUpdate(contentControl, content)
	}
}

func (this *DocumentEvents2DispImpl) BuildingBlockInsert(range_ *Range, name string, category string, blockType string, template string) {
	if this.Handlers.BuildingBlockInsert != nil {
		this.Handlers.BuildingBlockInsert(range_, name, category, blockType, template)
	}
}

type DocumentEvents2Impl struct {
	ole.IDispatchImpl
	DispImpl DocumentEvents2DispInterface
}

func (this *DocumentEvents2Impl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == IID_DocumentEvents2 {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IDispatchImpl.QueryInterface(riid, ppvObject)
}

func (this *DocumentEvents2Impl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32,
	wFlags uint16, pDispParams *win32.DISPPARAMS, pVarResult *win32.VARIANT,
	pExcepInfo *win32.EXCEPINFO, puArgErr *uint32) win32.HRESULT {
	var unwrapActions ole.Actions
	defer unwrapActions.Execute()
	switch dispIdMember {
	case 4:
		this.DispImpl.New()
		return win32.S_OK
	case 5:
		this.DispImpl.Open()
		return win32.S_OK
	case 6:
		this.DispImpl.Close()
		return win32.S_OK
	case 7:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1, _ := vArgs[0].ToInt32()
		this.DispImpl.Sync(p1)
		return win32.S_OK
	case 8:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*XMLNode)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToBool()
		this.DispImpl.XMLAfterInsert(p1, p2)
		return win32.S_OK
	case 9:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 3)
		p1 := (*Range)(vArgs[0].ToPointer())
		p2 := (*XMLNode)(vArgs[1].ToPointer())
		p3, _ := vArgs[2].ToBool()
		this.DispImpl.XMLBeforeDelete(p1, p2, p3)
		return win32.S_OK
	case 12:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*ContentControl)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToBool()
		this.DispImpl.ContentControlAfterAdd(p1, p2)
		return win32.S_OK
	case 13:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*ContentControl)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToBool()
		this.DispImpl.ContentControlBeforeDelete(p1, p2)
		return win32.S_OK
	case 14:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*ContentControl)(vArgs[0].ToPointer())
		p2 := (*win32.VARIANT_BOOL)(vArgs[1].ToPointer())
		this.DispImpl.ContentControlOnExit(p1, p2)
		return win32.S_OK
	case 15:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 1)
		p1 := (*ContentControl)(vArgs[0].ToPointer())
		this.DispImpl.ContentControlOnEnter(p1)
		return win32.S_OK
	case 16:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*ContentControl)(vArgs[0].ToPointer())
		p2 := (*win32.BSTR)(vArgs[1].ToPointer())
		this.DispImpl.ContentControlBeforeStoreUpdate(p1, p2)
		return win32.S_OK
	case 17:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 2)
		p1 := (*ContentControl)(vArgs[0].ToPointer())
		p2 := (*win32.BSTR)(vArgs[1].ToPointer())
		this.DispImpl.ContentControlBeforeContentUpdate(p1, p2)
		return win32.S_OK
	case 18:
		vArgs, _ := ole.ProcessInvokeArgs(pDispParams, 5)
		p1 := (*Range)(vArgs[0].ToPointer())
		p2, _ := vArgs[1].ToString()
		p3, _ := vArgs[2].ToString()
		p4, _ := vArgs[3].ToString()
		p5, _ := vArgs[4].ToString()
		this.DispImpl.BuildingBlockInsert(p1, p2, p3, p4, p5)
		return win32.S_OK
	}
	return win32.E_NOTIMPL
}

type DocumentEvents2ComObj struct {
	ole.IDispatchComObj
}

func NewDocumentEvents2ComObj(dispImpl DocumentEvents2DispInterface, scoped bool) *DocumentEvents2ComObj {
	comObj := com.NewComObj[DocumentEvents2ComObj](
		&DocumentEvents2Impl {DispImpl: dispImpl})
	if scoped {
		com.AddToScope(comObj)
	}
	return comObj
}

