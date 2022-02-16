#include "web_browser.h"
#include "doc_host_ui_handler.h"

#include <ExDispid.h>
#include <MsHTML.h>
#include <Shlwapi.h>
#pragma comment(lib, "Shlwapi.lib")

WebBrowser::WebBrowser(HWND hwndParent) : hwndParent_{hwndParent}
{
    if (CreateBrowser() == FALSE) {
        return;
    }

    ShowWindow(GetControlWindow(), SW_SHOW);
}

WebBrowser::~WebBrowser()
{
    delete docHostUiHandler_;
}

bool WebBrowser::CreateBrowser()
{
    HRESULT result;
    result = ::OleCreate(CLSID_WebBrowser, IID_IOleObject, OLERENDER_DRAW, 0, this, this, (void**)&oleObject_);
    if (SUCCEEDED(result) && oleObject_) {
        if (SUCCEEDED(result)) {
            result = oleObject_->SetClientSite(this);
        }
        if (SUCCEEDED(result)) {
            result = OleSetContainedObject(oleObject_, TRUE);
        }

        if (SUCCEEDED(result)) {
            RECT posRect;
            ::SetRect(&posRect, 0, 0, 300, 300);
            result = oleObject_->DoVerb(OLEIVERB_INPLACEACTIVATE, NULL, this, -1, hwndParent_, &posRect);
        }
        result = oleObject_->QueryInterface(&webBrowser_);

        IConnectionPointContainer* container = nullptr;
        IUnknown* punk = nullptr;
        result = webBrowser_->QueryInterface(IID_IConnectionPointContainer, (void**)&container);
        if (SUCCEEDED(result) && container) {
            result = container->FindConnectionPoint(__uuidof(DWebBrowserEvents2), &callback_);
        }
        if (SUCCEEDED(result) && container) {
            result = this->QueryInterface(IID_IUnknown, (void**)&punk);
        }
        if (SUCCEEDED(result) && container) {
            result = callback_->Advise(punk, &eventCookie_);
        }

        if (docHostUiHandler_) {
            delete docHostUiHandler_;
        }

        docHostUiHandler_ = new DocHostUIHandler{};
    }

    if (FAILED(result) || !webBrowser_) {
        return false;
    }
    return true;
}

void WebBrowser::SetHtml(const StringType& html)
{
    html_ = html;
    Navigate(L"about:blank");
}

RECT WebBrowser::PixelToHiMetric(const RECT& rect)
{
    static bool initialized = false;
    static int pixelsPerInchX = 0;
    static int pixelsPerInchY = 0;
    if (!initialized) {
        HDC hdc = ::GetDC(0);
        pixelsPerInchX = ::GetDeviceCaps(hdc, LOGPIXELSX);
        pixelsPerInchY = ::GetDeviceCaps(hdc, LOGPIXELSY);
        ::ReleaseDC(0, hdc);
        initialized = true;
    }

    RECT rc;
    rc.left = MulDiv(2540, rect.left, pixelsPerInchX);
    rc.top = MulDiv(2540, rect.top, pixelsPerInchY);
    rc.right = MulDiv(2540, rect.right, pixelsPerInchX);
    rc.bottom = MulDiv(2540, rect.bottom, pixelsPerInchY);
    return rc;
}

void WebBrowser::SetRect(const RECT& rect)
{
    browserRect_ = rect;

    RECT hiMetricRect = PixelToHiMetric(browserRect_);
    SIZEL sz;
    sz.cx = hiMetricRect.right - hiMetricRect.left;
    sz.cy = hiMetricRect.bottom - hiMetricRect.top;
    oleObject_->SetExtent(DVASPECT_CONTENT, &sz);

    if (oleInPlaceObject_ != 0) {
        oleInPlaceObject_->SetObjectRects(&browserRect_, &browserRect_);
    }
}

bool WebBrowser::GetCheckboxValue(const std::wstring& checkboxId, bool fallbackValue)
{
    IDispatch* dispatch;
    auto result = webBrowser_->get_Document(&dispatch);
    if (FAILED(result) || !dispatch) {
        return fallbackValue;
    }

    IHTMLDocument3* document;
    result = dispatch->QueryInterface<IHTMLDocument3>(&document);
    if (FAILED(result) || !document) {
        return fallbackValue;
    }

    IHTMLElement* element;
    BSTR urlBstr = ::SysAllocString(checkboxId.c_str());
    result = document->getElementById(urlBstr, &element);
    if (FAILED(result) || !element) {
        return fallbackValue;
    }

    BSTR attribute = ::SysAllocString(L"checked");
    VARIANT valueVariant;
    result = element->getAttribute(attribute, 0, &valueVariant);
    if (SUCCEEDED(result)) {
        return (valueVariant.boolVal != 0);
    }
    else {
        return fallbackValue;
    }
}

void WebBrowser::SetLinkClickCallback(const std::function<void(const std::wstring&)>& callback)
{
    linkClickCallback_ = callback;
}

// ----- Control methods -----

void WebBrowser::GoBack()
{
    webBrowser_->GoBack();
}

void WebBrowser::GoForward()
{
    webBrowser_->GoForward();
}

void WebBrowser::Refresh()
{
    webBrowser_->Refresh();
}

void WebBrowser::Navigate(std::wstring url)
{
    bstr_t burl{url.c_str()};
    variant_t flags{navNoHistory};
    webBrowser_->Navigate(burl, &flags, 0, 0, 0);
}

// ----- IUnknown -----

HRESULT STDMETHODCALLTYPE WebBrowser::QueryInterface(REFIID riid, void** ppvObject)
{
    if (riid == __uuidof(IUnknown)) {
        *ppvObject = static_cast<IDispatch*>(this);
    }
    else if (riid == __uuidof(IDispatch)) {
        *ppvObject = static_cast<IDispatch*>(this);
    }
    else if (riid == __uuidof(IOleClientSite)) {
        *ppvObject = static_cast<IOleClientSite*>(this);
    }
    else if (riid == __uuidof(IOleInPlaceSite)) {
        *ppvObject = static_cast<IOleInPlaceSite*>(this);
    }
    else if (riid == __uuidof(IDocHostUIHandler)) {
        *ppvObject = docHostUiHandler_;
    }
    else {
        return E_NOINTERFACE;
    }

    AddRef();

    return S_OK;
}

ULONG STDMETHODCALLTYPE WebBrowser::AddRef(void)
{
    comRefCount_++;
    return comRefCount_;
}

ULONG STDMETHODCALLTYPE WebBrowser::Release(void)
{
    comRefCount_--;
    return comRefCount_;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetTypeInfoCount(UINT* pctinfo)
{
    return E_FAIL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetTypeInfo(UINT, LCID, ITypeInfo**)
{
    return E_FAIL;
}

HRESULT STDMETHODCALLTYPE
    WebBrowser::GetIDsOfNames(REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
    return E_FAIL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Invoke(
    DISPID dispIdMember,
    REFIID riid,
    LCID lcid,
    WORD wFlags,
    DISPPARAMS* pDispParams,
    VARIANT* pVarResult,
    EXCEPINFO* pExcepInfo,
    UINT* puArgErr)
{
    if (dispIdMember == DISPID_DOCUMENTCOMPLETE) {
        return OnDocumentCompleted(pDispParams);
    }
    else if (dispIdMember == DISPID_BEFORENAVIGATE2) {
        BSTR url = pDispParams->rgvarg[DispArgumentIds::Url].pvarVal->bstrVal;
        if (wcscmp(url, L"about:blank") != 0) {
            *(pDispParams->rgvarg[DispArgumentIds::CancelNavigation].pboolVal) = TRUE;
            std::wstring urlString{url, SysStringLen(url)};
            if (linkClickCallback_) {
                linkClickCallback_(urlString);
            }
        }
    }
    
    return S_OK;
}

// ---------- IOleWindow ----------

HRESULT STDMETHODCALLTYPE WebBrowser::GetWindow(__RPC__deref_out_opt HWND* phwnd)
{
    (*phwnd) = hwndParent_;
    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::ContextSensitiveHelp(BOOL fEnterMode)
{
    return E_NOTIMPL;
}

// ---------- IOleInPlaceSite ----------

HRESULT STDMETHODCALLTYPE WebBrowser::CanInPlaceActivate(void)
{
    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnInPlaceActivate(void)
{
    OleLockRunning(oleObject_, TRUE, FALSE);
    oleObject_->QueryInterface(&oleInPlaceObject_);
    oleInPlaceObject_->SetObjectRects(&browserRect_, &browserRect_);

    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnUIActivate(void)
{
    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetWindowContext(
    __RPC__deref_out_opt IOleInPlaceFrame** ppFrame,
    __RPC__deref_out_opt IOleInPlaceUIWindow** ppDoc,
    __RPC__out LPRECT lprcPosRect,
    __RPC__out LPRECT lprcClipRect,
    __RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo)
{
    HWND hwnd = hwndParent_;

    (*ppFrame) = NULL;
    (*ppDoc) = NULL;
    (*lprcPosRect).left = browserRect_.left;
    (*lprcPosRect).top = browserRect_.top;
    (*lprcPosRect).right = browserRect_.right;
    (*lprcPosRect).bottom = browserRect_.bottom;
    *lprcClipRect = *lprcPosRect;

    lpFrameInfo->fMDIApp = false;
    lpFrameInfo->hwndFrame = hwnd;
    lpFrameInfo->haccel = NULL;
    lpFrameInfo->cAccelEntries = 0;

    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Scroll(SIZE scrollExtant)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnUIDeactivate(BOOL fUndoable)
{
    return S_OK;
}

HWND WebBrowser::GetControlWindow()
{
    if (hwndControl_ != 0)
        return hwndControl_;

    if (oleInPlaceObject_ == 0)
        return 0;

    oleInPlaceObject_->GetWindow(&hwndControl_);
    return hwndControl_;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnInPlaceDeactivate(void)
{
    hwndControl_ = 0;
    oleInPlaceObject_ = 0;

    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::DiscardUndoState(void)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::DeactivateAndUndo(void)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnPosRectChange(__RPC__in LPCRECT lprcPosRect)
{
    return E_NOTIMPL;
}

// ---------- IOleClientSite ----------

HRESULT STDMETHODCALLTYPE WebBrowser::SaveObject(void)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE
    WebBrowser::GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, __RPC__deref_out_opt IMoniker** ppmk)
{
    if ((dwAssign == OLEGETMONIKER_ONLYIFTHERE) && (dwWhichMoniker == OLEWHICHMK_CONTAINER))
        return E_FAIL;

    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::GetContainer(__RPC__deref_out_opt IOleContainer** ppContainer)
{
    return E_NOINTERFACE;
}

HRESULT STDMETHODCALLTYPE WebBrowser::ShowObject(void)
{
    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OnShowWindow(BOOL fShow)
{
    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::RequestNewObjectLayout(void)
{
    return E_NOTIMPL;
}

// ----- IStorage -----

HRESULT STDMETHODCALLTYPE WebBrowser::CreateStream(
    __RPC__in_string const OLECHAR* pwcsName,
    DWORD grfMode,
    DWORD reserved1,
    DWORD reserved2,
    __RPC__deref_out_opt IStream** ppstm)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE
    WebBrowser::OpenStream(const OLECHAR* pwcsName, void* reserved1, DWORD grfMode, DWORD reserved2, IStream** ppstm)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::CreateStorage(
    __RPC__in_string const OLECHAR* pwcsName,
    DWORD grfMode,
    DWORD reserved1,
    DWORD reserved2,
    __RPC__deref_out_opt IStorage** ppstg)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::OpenStorage(
    __RPC__in_opt_string const OLECHAR* pwcsName,
    __RPC__in_opt IStorage* pstgPriority,
    DWORD grfMode,
    __RPC__deref_opt_in_opt SNB snbExclude,
    DWORD reserved,
    __RPC__deref_out_opt IStorage** ppstg)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE
    WebBrowser::CopyTo(DWORD ciidExclude, const IID* rgiidExclude, __RPC__in_opt SNB snbExclude, IStorage* pstgDest)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::MoveElementTo(
    __RPC__in_string const OLECHAR* pwcsName,
    __RPC__in_opt IStorage* pstgDest,
    __RPC__in_string const OLECHAR* pwcsNewName,
    DWORD grfFlags)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Commit(DWORD grfCommitFlags)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Revert(void)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE
    WebBrowser::EnumElements(DWORD reserved1, void* reserved2, DWORD reserved3, IEnumSTATSTG** ppenum)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::DestroyElement(__RPC__in_string const OLECHAR* pwcsName)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE
    WebBrowser::RenameElement(__RPC__in_string const OLECHAR* pwcsOldName, __RPC__in_string const OLECHAR* pwcsNewName)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::SetElementTimes(
    __RPC__in_opt_string const OLECHAR* pwcsName,
    __RPC__in_opt const FILETIME* pctime,
    __RPC__in_opt const FILETIME* patime,
    __RPC__in_opt const FILETIME* pmtime)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::SetClass(__RPC__in REFCLSID clsid)
{
    return S_OK;
}

HRESULT STDMETHODCALLTYPE WebBrowser::SetStateBits(DWORD grfStateBits, DWORD grfMask)
{
    return E_NOTIMPL;
}

HRESULT STDMETHODCALLTYPE WebBrowser::Stat(__RPC__out STATSTG* pstatstg, DWORD grfStatFlag)
{
    return E_NOTIMPL;
}

static int count = 0;

HRESULT WebBrowser::OnDocumentCompleted(DISPPARAMS* args)
{
    if (html_.empty()) {
        return S_OK;
    }

    IDispatch* iDispatch = 0;
    IHTMLDocument2* htmlDocument = 0;
    HRESULT result = webBrowser_->get_Document(&iDispatch);
    if (SUCCEEDED(result) && iDispatch) {
        result = iDispatch->QueryInterface(IID_IHTMLDocument2, (void**)&htmlDocument);
    }

    IStream* iStream = 0;
    IPersistStreamInit* htmlDocumentPersistStream = 0;
    if (SUCCEEDED(result) && htmlDocument) {
        result = htmlDocument->QueryInterface(IID_IPersistStreamInit, (void**)&htmlDocumentPersistStream);
    }

    if (SUCCEEDED(result) && htmlDocumentPersistStream) {
        //iStream = SHCreateMemStream((BYTE*)html_.c_str(), html_.size() * sizeof(wchar_t));
        iStream = SHCreateMemStream((BYTE*)html_.c_str(), html_.size() * sizeof(StringType::value_type));
        result = htmlDocumentPersistStream->Load(iStream);
    }

    if (iStream) {
        iStream->Release();
    }
    if (htmlDocumentPersistStream) {
        htmlDocumentPersistStream->Release();
    }
    if (htmlDocument) {
        htmlDocument->Release();
    }
    if (iDispatch) {
        iDispatch->Release();
    }

    html_.clear();

    return S_OK;
}
