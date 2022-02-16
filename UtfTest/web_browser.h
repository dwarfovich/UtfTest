#include <comdef.h>
#include <Exdisp.h>

#include <string>
#include <functional>

class DocHostUIHandler;

using StringType = std::string;

class WebBrowser : public IDispatch, public IOleClientSite, public IOleInPlaceSite, public IStorage {
public:
    WebBrowser(HWND hwndParent);
    ~WebBrowser();

    void SetHtml(const StringType& html);
    RECT PixelToHiMetric(const RECT& rect);
    void SetRect(const RECT& rect);
    bool GetCheckboxValue(const std::wstring& checkboxId, bool fallbackValue = true);
    void SetLinkClickCallback(const std::function<void(const std::wstring& url)>& callback);

    // ----- Control methods -----
    void GoBack();
    void GoForward();
    void Refresh();
    void Navigate(std::wstring url);

    // ----- IUnknown -----
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override;
    ULONG STDMETHODCALLTYPE AddRef(void) override;
    ULONG STDMETHODCALLTYPE Release(void) override;

    // ---------- IDispatch ----------
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(__RPC__out UINT* pctinfo) override;
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, __RPC__deref_out_opt ITypeInfo**) override;
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(
        __RPC__in REFIID riid,
        __RPC__in_ecount_full(cNames) LPOLESTR* rgszNames,
        __RPC__in_range(0, 16384) UINT cNames,
        LCID lcid,
        __RPC__out_ecount_full(cNames) DISPID* rgDispId) override;
    HRESULT STDMETHODCALLTYPE Invoke(
        _In_ DISPID dispIdMember,
        _In_ REFIID riid,
        _In_ LCID lcid,
        _In_ WORD wFlags,
        _In_ DISPPARAMS* pDispParams,
        _Out_opt_ VARIANT* pVarResult,
        _Out_opt_ EXCEPINFO* pExcepInfo,
        _Out_opt_ UINT* puArgErr) override;

    // ---------- IOleWindow ----------
    HRESULT STDMETHODCALLTYPE GetWindow(__RPC__deref_out_opt HWND* phwnd) override;
    HRESULT STDMETHODCALLTYPE ContextSensitiveHelp(BOOL fEnterMode) override;

    // ---------- IOleInPlaceSite ----------
    HRESULT STDMETHODCALLTYPE CanInPlaceActivate(void) override;
    HRESULT STDMETHODCALLTYPE OnInPlaceActivate(void) override;
    HRESULT STDMETHODCALLTYPE OnUIActivate(void) override;
    HRESULT STDMETHODCALLTYPE GetWindowContext(
        __RPC__deref_out_opt IOleInPlaceFrame** ppFrame,
        __RPC__deref_out_opt IOleInPlaceUIWindow** ppDoc,
        __RPC__out LPRECT lprcPosRect,
        __RPC__out LPRECT lprcClipRect,
        __RPC__inout LPOLEINPLACEFRAMEINFO lpFrameInfo) override;
    HRESULT STDMETHODCALLTYPE Scroll(SIZE scrollExtant) override;
    HRESULT STDMETHODCALLTYPE OnUIDeactivate(BOOL fUndoable) override;
    HWND GetControlWindow();
    HRESULT STDMETHODCALLTYPE OnInPlaceDeactivate(void) override;
    HRESULT STDMETHODCALLTYPE DiscardUndoState(void) override;
    HRESULT STDMETHODCALLTYPE DeactivateAndUndo(void) override;
    HRESULT STDMETHODCALLTYPE OnPosRectChange(__RPC__in LPCRECT lprcPosRect) override;

    // ---------- IOleClientSite ----------
    HRESULT STDMETHODCALLTYPE SaveObject(void) override;
    HRESULT STDMETHODCALLTYPE
        GetMoniker(DWORD dwAssign, DWORD dwWhichMoniker, __RPC__deref_out_opt IMoniker** ppmk) override;
    HRESULT STDMETHODCALLTYPE GetContainer(__RPC__deref_out_opt IOleContainer** ppContainer) override;
    HRESULT STDMETHODCALLTYPE ShowObject(void) override;
    HRESULT STDMETHODCALLTYPE OnShowWindow(BOOL fShow) override;
    HRESULT STDMETHODCALLTYPE RequestNewObjectLayout(void) override;

    // ----- IStorage -----
    HRESULT STDMETHODCALLTYPE CreateStream(
        __RPC__in_string const OLECHAR* pwcsName,
        DWORD grfMode,
        DWORD reserved1,
        DWORD reserved2,
        __RPC__deref_out_opt IStream** ppstm) override;
    HRESULT STDMETHODCALLTYPE
        OpenStream(const OLECHAR* pwcsName, void* reserved1, DWORD grfMode, DWORD reserved2, IStream** ppstm) override;
    HRESULT STDMETHODCALLTYPE CreateStorage(
        __RPC__in_string const OLECHAR* pwcsName,
        DWORD grfMode,
        DWORD reserved1,
        DWORD reserved2,
        __RPC__deref_out_opt IStorage** ppstg) override;
    HRESULT STDMETHODCALLTYPE OpenStorage(
        __RPC__in_opt_string const OLECHAR* pwcsName,
        __RPC__in_opt IStorage* pstgPriority,
        DWORD grfMode,
        __RPC__deref_opt_in_opt SNB snbExclude,
        DWORD reserved,
        __RPC__deref_out_opt IStorage** ppstg) override;
    HRESULT STDMETHODCALLTYPE
        CopyTo(DWORD ciidExclude, const IID* rgiidExclude, __RPC__in_opt SNB snbExclude, IStorage* pstgDest) override;
    HRESULT STDMETHODCALLTYPE MoveElementTo(
        __RPC__in_string const OLECHAR* pwcsName,
        __RPC__in_opt IStorage* pstgDest,
        __RPC__in_string const OLECHAR* pwcsNewName,
        DWORD grfFlags) override;
    HRESULT STDMETHODCALLTYPE Commit(DWORD grfCommitFlags) override;
    HRESULT STDMETHODCALLTYPE Revert(void) override;
    HRESULT STDMETHODCALLTYPE
        EnumElements(DWORD reserved1, void* reserved2, DWORD reserved3, IEnumSTATSTG** ppenum) override;
    HRESULT STDMETHODCALLTYPE DestroyElement(__RPC__in_string const OLECHAR* pwcsName) override;
    HRESULT STDMETHODCALLTYPE RenameElement(
        __RPC__in_string const OLECHAR* pwcsOldName,
        __RPC__in_string const OLECHAR* pwcsNewName) override;
    HRESULT STDMETHODCALLTYPE SetElementTimes(
        __RPC__in_opt_string const OLECHAR* pwcsName,
        __RPC__in_opt const FILETIME* pctime,
        __RPC__in_opt const FILETIME* patime,
        __RPC__in_opt const FILETIME* pmtime) override;
    HRESULT STDMETHODCALLTYPE SetClass(__RPC__in REFCLSID clsid) override;
    HRESULT STDMETHODCALLTYPE SetStateBits(DWORD grfStateBits, DWORD grfMask) override;
    HRESULT STDMETHODCALLTYPE Stat(__RPC__out STATSTG* pstatstg, DWORD grfStatFlag) override;

protected:  // methods
    bool CreateBrowser();
    HRESULT OnDocumentCompleted(DISPPARAMS* args);

protected:  // data
    enum DispArgumentIds { Url = 5, CancelNavigation = 0 };

    HWND hwndParent_ = NULL;
    HWND hwndControl_ = NULL;
    RECT browserRect_;

    IOleObject* oleObject_ = nullptr;
    IOleInPlaceObject* oleInPlaceObject_ = nullptr;
    IWebBrowser2* webBrowser_ = nullptr;
    DocHostUIHandler* docHostUiHandler_ = nullptr;
    IConnectionPoint* callback_ = nullptr;
    LONG comRefCount_ = 0;
    DWORD eventCookie_ = 0;

    StringType                                   html_;
    std::function<void(const std::wstring& url)> linkClickCallback_;
};