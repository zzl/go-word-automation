package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
	"unsafe"
)

// 000209F6-0000-0000-C000-000000000046
var IID_DocumentEvents = syscall.GUID{0x000209F6, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type DocumentEventsDispInterface interface {
	New() 
	Open() 
	Close() 
}

type DocumentEventsHandlers struct {
	New func() 
	Open func() 
	Close func() 
}

type DocumentEventsDispImpl struct {
	Handlers DocumentEventsHandlers
}

func (this *DocumentEventsDispImpl) New() {
	if this.Handlers.New != nil {
		this.Handlers.New()
	}
}

func (this *DocumentEventsDispImpl) Open() {
	if this.Handlers.Open != nil {
		this.Handlers.Open()
	}
}

func (this *DocumentEventsDispImpl) Close() {
	if this.Handlers.Close != nil {
		this.Handlers.Close()
	}
}

type DocumentEventsImpl struct {
	ole.IDispatchImpl
	DispImpl DocumentEventsDispInterface
}

func (this *DocumentEventsImpl) QueryInterface(riid *syscall.GUID, ppvObject unsafe.Pointer) win32.HRESULT {
	if *riid == IID_DocumentEvents {
		this.AssignPpvObject(ppvObject)
		this.AddRef()
		return win32.S_OK
	}
	return this.IDispatchImpl.QueryInterface(riid, ppvObject)
}

func (this *DocumentEventsImpl) Invoke(dispIdMember int32, riid *syscall.GUID, lcid uint32,
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
	}
	return win32.E_NOTIMPL
}

type DocumentEventsComObj struct {
	ole.IDispatchComObj
}

func NewDocumentEventsComObj(dispImpl DocumentEventsDispInterface, scoped bool) *DocumentEventsComObj {
	comObj := com.NewComObj[DocumentEventsComObj](
		&DocumentEventsImpl {DispImpl: dispImpl})
	if scoped {
		com.AddToScope(comObj)
	}
	return comObj
}

