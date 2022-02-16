#pragma once

#include <MsHtmHst.h>

class DocHostUIHandler : public IDocHostUIHandler
{
    // Inherited via IDocHostUIHandler
    virtual HRESULT __stdcall QueryInterface(REFIID riid, void** ppvObject) override;
    virtual ULONG __stdcall AddRef(void) override;
    virtual ULONG __stdcall Release(void) override;
    virtual HRESULT __stdcall ShowContextMenu(DWORD      dwID,
                                              POINT*     ppt,
                                              IUnknown*  pcmdtReserved,
                                              IDispatch* pdispReserved) override;
    virtual HRESULT __stdcall GetHostInfo(DOCHOSTUIINFO* pInfo) override;
    virtual HRESULT __stdcall ShowUI(DWORD                    dwID,
                                     IOleInPlaceActiveObject* pActiveObject,
                                     IOleCommandTarget*       pCommandTarget,
                                     IOleInPlaceFrame*        pFrame,
                                     IOleInPlaceUIWindow*     pDoc) override;
    virtual HRESULT __stdcall HideUI(void) override;
    virtual HRESULT __stdcall UpdateUI(void) override;
    virtual HRESULT __stdcall EnableModeless(BOOL fEnable) override;
    virtual HRESULT __stdcall OnDocWindowActivate(BOOL fActivate) override;
    virtual HRESULT __stdcall OnFrameWindowActivate(BOOL fActivate) override;
    virtual HRESULT __stdcall ResizeBorder(LPCRECT              prcBorder,
                                           IOleInPlaceUIWindow* pUIWindow,
                                           BOOL                 fRameWindow) override;
    virtual HRESULT __stdcall GetOptionKeyPath(LPOLESTR* pchKey, DWORD dw) override;
    virtual HRESULT __stdcall GetDropTarget(IDropTarget* pDropTarget, IDropTarget** ppDropTarget) override;
    virtual HRESULT __stdcall GetExternal(IDispatch** ppDispatch) override;
    virtual HRESULT __stdcall TranslateUrl(DWORD dwTranslate, LPWSTR pchURLIn, LPWSTR* ppchURLOut) override;
    virtual HRESULT __stdcall FilterDataObject(IDataObject* pDO, IDataObject** ppDORet) override;
    virtual HRESULT __stdcall TranslateAcceleratorW(LPMSG lpMsg, const GUID* pguidCmdGroup, DWORD nCmdID);
};
