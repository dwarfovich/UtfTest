#include "doc_host_ui_handler.h"

HRESULT __stdcall DocHostUIHandler::QueryInterface(REFIID riid, void** ppvObject)
{
    return E_NOINTERFACE;
}

ULONG __stdcall DocHostUIHandler::AddRef(void)
{
    return 1;
}

ULONG __stdcall DocHostUIHandler::Release(void)
{
    return 1;
}

HRESULT __stdcall DocHostUIHandler::ShowContextMenu(DWORD      dwID,
                                                    POINT*     ppt,
                                                    IUnknown*  pcmdtReserved,
                                                    IDispatch* pdispReserved)
{
    return S_OK;
}

HRESULT __stdcall DocHostUIHandler::GetHostInfo(DOCHOSTUIINFO* pInfo)
{
    pInfo->cbSize = sizeof(DOCHOSTUIINFO);

    pInfo->dwFlags =
        //DOCHOSTUIFLAG_DPI_AWARE | DOCHOSTUIFLAG_NO3DBORDER | DOCHOSTUIFLAG_SCROLL_NO | DOCHOSTUIFLAG_NO3DOUTERBORDER;
        DOCHOSTUIFLAG_DPI_AWARE | DOCHOSTUIFLAG_SCROLL_NO;
    pInfo->dwDoubleClick = DOCHOSTUIDBLCLK_DEFAULT;

    return S_OK;
}

HRESULT __stdcall DocHostUIHandler::ShowUI(DWORD                    dwID,
                                           IOleInPlaceActiveObject* pActiveObject,
                                           IOleCommandTarget*       pCommandTarget,
                                           IOleInPlaceFrame*        pFrame,
                                           IOleInPlaceUIWindow*     pDoc)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::HideUI(void)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::UpdateUI(void)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::EnableModeless(BOOL fEnable)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::OnDocWindowActivate(BOOL fActivate)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::OnFrameWindowActivate(BOOL fActivate)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::ResizeBorder(LPCRECT prcBorder, IOleInPlaceUIWindow* pUIWindow, BOOL fRameWindow)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::GetOptionKeyPath(LPOLESTR* pchKey, DWORD dw)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::GetDropTarget(IDropTarget* pDropTarget, IDropTarget** ppDropTarget)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::GetExternal(IDispatch** ppDispatch)
{
    *ppDispatch = 0;
    return S_FALSE;
}

HRESULT __stdcall DocHostUIHandler::TranslateUrl(DWORD dwTranslate, LPWSTR pchURLIn, LPWSTR* ppchURLOut)
{
    return E_NOTIMPL;
}

HRESULT __stdcall DocHostUIHandler::FilterDataObject(IDataObject* pDO, IDataObject** ppDORet)
{
    return E_NOTIMPL;
}

HRESULT DocHostUIHandler::TranslateAcceleratorW(LPMSG lpMsg, const GUID* pguidCmdGroup, DWORD nCmdID)
{
    return E_NOTIMPL;
}
