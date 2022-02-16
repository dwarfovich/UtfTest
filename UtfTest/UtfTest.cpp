#include "framework.h"
#include "UtfTest.h"
#include "web_browser.h"

#include <string>

#define MAX_LOADSTRING 100

HINSTANCE   hInst;                         // current instance
WCHAR       szTitle[MAX_LOADSTRING];       // The title bar text
WCHAR       szWindowClass[MAX_LOADSTRING]; // the main window class name
HWND        mainWindowHwnd = NULL;
WebBrowser* webBrowser     = nullptr;

ATOM             MyRegisterClass(HINSTANCE hInstance);
BOOL             InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK About(HWND, UINT, WPARAM, LPARAM);

static const std::string html = R"(
<!DOCTYPE html> 
<head>
    <meta charset="UTF-8">
</head>
<body>
    <p>Привет</p>
</body>
</html>
)";

int APIENTRY wWinMain(_In_ HINSTANCE     hInstance,
                      _In_opt_ HINSTANCE hPrevInstance,
                      _In_ LPWSTR        lpCmdLine,
                      _In_ int           nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_UTFTEST, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance(hInstance, nCmdShow)) {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_UTFTEST));

    MSG msg;

    std::wstring text          = L"Привет";
    HDC          mainWindowHdc = GetDC(mainWindowHwnd);
    RECT         textRect      = { 10, 50, 200, 100 };
    DrawText(mainWindowHdc, text.c_str(), -1, &textRect, 0);

    webBrowser = new WebBrowser(mainWindowHwnd);

    RECT browserRect { 10, 200, 400, 400 };

    SetWindowPos(webBrowser->GetControlWindow(),
                 HWND_TOP,
                 browserRect.left,
                 browserRect.top,
                 browserRect.left + 300,
                 browserRect.top + 200,
                 0);

    webBrowser->SetHtml(html);

    while (GetMessage(&msg, nullptr, 0, 0)) {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg)) {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int)msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style         = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc   = WndProc;
    wcex.cbClsExtra    = 0;
    wcex.cbWndExtra    = 0;
    wcex.hInstance     = hInstance;
    wcex.hIcon         = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_UTFTEST));
    wcex.hCursor       = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszMenuName  = MAKEINTRESOURCEW(IDC_UTFTEST);
    wcex.lpszClassName = szWindowClass;
    wcex.hIconSm       = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    hInst = hInstance; // Store instance handle in our global variable

    mainWindowHwnd = CreateWindowW(szWindowClass,
                                   szTitle,
                                   WS_OVERLAPPEDWINDOW,
                                   CW_USEDEFAULT,
                                   0,
                                   CW_USEDEFAULT,
                                   0,
                                   nullptr,
                                   nullptr,
                                   hInstance,
                                   nullptr);

    if (!mainWindowHwnd) {
        return FALSE;
    }

    ShowWindow(mainWindowHwnd, nCmdShow);
    UpdateWindow(mainWindowHwnd);

    return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message) {
        case WM_COMMAND: {
            int wmId = LOWORD(wParam);
            switch (wmId) {
                case IDM_EXIT: DestroyWindow(hWnd); break;
                default: return DefWindowProc(hWnd, message, wParam, lParam);
            }
        } break;
        case WM_PAINT: {
            PAINTSTRUCT ps;
            HDC         hdc = BeginPaint(hWnd, &ps);
            EndPaint(hWnd, &ps);
        } break;
        case WM_DESTROY: PostQuitMessage(0); break;
        default: return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}