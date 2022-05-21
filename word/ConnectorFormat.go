package word

import (
	"github.com/zzl/go-win32api/win32"
	"github.com/zzl/go-com/com"
	"github.com/zzl/go-com/ole"
	"syscall"
)

// 000209C7-0000-0000-C000-000000000046
var IID_ConnectorFormat = syscall.GUID{0x000209C7, 0x0000, 0x0000, 
	[8]byte{0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}}

type ConnectorFormat struct {
	ole.OleClient
}

func NewConnectorFormat(pDisp *win32.IDispatch, addRef bool, scoped bool) *ConnectorFormat {
	p := &ConnectorFormat{ole.OleClient{pDisp}}
	if addRef {
		pDisp.AddRef()
	}
	if scoped {
		com.AddToScope(p)
	}
	return p
}

func ConnectorFormatFromVar(v ole.Variant) *ConnectorFormat {
	return NewConnectorFormat(v.PdispValVal(), false, false)
}

func (this *ConnectorFormat) IID() *syscall.GUID {
	return &IID_ConnectorFormat
}

func (this *ConnectorFormat) GetIDispatch(addRef bool) *win32.IDispatch {
	if addRef {
		this.AddRef()
	}
	return this.IDispatch
}

func (this *ConnectorFormat) Application() *Application {
	retVal := this.PropGet(0x000003e8, nil)
	return NewApplication(retVal.PdispValVal(), false, true)
}

func (this *ConnectorFormat) Creator() int32 {
	retVal := this.PropGet(0x000003e9, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) BeginConnected() int32 {
	retVal := this.PropGet(0x00000064, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) BeginConnectedShape() *Shape {
	retVal := this.PropGet(0x00000065, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *ConnectorFormat) BeginConnectionSite() int32 {
	retVal := this.PropGet(0x00000066, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) EndConnected() int32 {
	retVal := this.PropGet(0x00000067, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) EndConnectedShape() *Shape {
	retVal := this.PropGet(0x00000068, nil)
	return NewShape(retVal.PdispValVal(), false, true)
}

func (this *ConnectorFormat) EndConnectionSite() int32 {
	retVal := this.PropGet(0x00000069, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) Parent() *ole.DispatchClass {
	retVal := this.PropGet(0x00000001, nil)
	return ole.NewDispatchClass(retVal.PdispValVal(), true)
}

func (this *ConnectorFormat) Type() int32 {
	retVal := this.PropGet(0x0000006a, nil)
	return retVal.LValVal()
}

func (this *ConnectorFormat) SetType(rhs int32)  {
	retVal := this.PropPut(0x0000006a, []interface{}{rhs})
	_= retVal
}

func (this *ConnectorFormat) BeginConnect(connectedShape **Shape, connectionSite int32)  {
	retVal := this.Call(0x0000000a, []interface{}{connectedShape, connectionSite})
	_= retVal
}

func (this *ConnectorFormat) BeginDisconnect()  {
	retVal := this.Call(0x0000000b, nil)
	_= retVal
}

func (this *ConnectorFormat) EndConnect(connectedShape **Shape, connectionSite int32)  {
	retVal := this.Call(0x0000000c, []interface{}{connectedShape, connectionSite})
	_= retVal
}

func (this *ConnectorFormat) EndDisconnect()  {
	retVal := this.Call(0x0000000d, nil)
	_= retVal
}

