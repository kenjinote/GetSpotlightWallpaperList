#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "wininet")
#pragma comment(lib, "Shlwapi")
#pragma comment(lib, "gdiplus")
#pragma comment (lib, "shlwapi")

#include <windows.h>
#include <wininet.h>
#include <shlwapi.h>
#include <gdiplus.h>
#include <shlobj.h>
#include <shlwapi.h>

using namespace Gdiplus;

TCHAR szClassName[] = TEXT("Window");

WNDPROC ListWndProc;
#define WM_CALCPARAMETER (WM_APP)
#define WM_INVALIDATEITEM (WM_APP+1)
#define WM_ENSUREVISIBLE (WM_APP+2)
#define WM_CREATED (WM_APP+3)
#define MAX_TEXT_LENGTH (256)
#define DRAG_IMAGE_WIDTH (238 / 2)
#define DRAG_IMAGE_HEIGHT (240 / 2)

typedef struct
{
	Gdiplus::Image *image;
	TCHAR szImageFilePath[MAX_PATH];
} LIST_ITEM;

BOOL SetDragImage(IN HWND hWnd, IN IDragSourceHelper *pDragSourceHelper, IN IDataObject *pDataObject, IN Gdiplus::Image *image, IN LPPOINT pt)
{
	SIZE size = { DRAG_IMAGE_WIDTH, DRAG_IMAGE_HEIGHT };
	HDC hdc = GetDC(hWnd);
	HDC hdcMem = CreateCompatibleDC(hdc);
	HBITMAP hbmp = CreateCompatibleBitmap(hdc, size.cx, size.cy);
	HBITMAP hbmpPrev = (HBITMAP)SelectObject(hdcMem, hbmp);
	{
		Gdiplus::Graphics g(hdcMem);
		g.DrawImage(image, 0, 0, size.cx, size.cy);
	}
	SelectObject(hdcMem, hbmpPrev);
	DeleteDC(hdcMem);
	ReleaseDC(hWnd, hdc);
	SHDRAGIMAGE dragImage;
	dragImage.sizeDragImage = size;
	dragImage.ptOffset = *pt;
	dragImage.hbmpDragImage = hbmp;
	dragImage.crColorKey = -1;
	HRESULT hr = pDragSourceHelper->InitializeFromBitmap(&dragImage, pDataObject);
	return hr == S_OK;
}

class CDropTarget : public IDropTarget
{
public:
	CDropTarget(HWND hwnd)
	{
		m_cRef = 1;
		m_hwnd = hwnd;
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_pDropTargetHelper));
	}
	~CDropTarget()
	{
		if (m_pDropTargetHelper != NULL)
			m_pDropTargetHelper->Release();
	}
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject)
	{
		*ppvObject = NULL;
		if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDropTarget))
			*ppvObject = static_cast<IDropTarget *>(this);
		else
			return E_NOINTERFACE;
		AddRef();
		return S_OK;
	}
	STDMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&m_cRef);
	}
	STDMETHODIMP_(ULONG) Release()
	{
		if (InterlockedDecrement(&m_cRef) == 0)
		{
			delete this;
			return 0;
		}
		return m_cRef;
	}
	STDMETHODIMP DragEnter(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
	{
		if (m_pDropTargetHelper != NULL)
			m_pDropTargetHelper->DragEnter(m_hwnd, pDataObj, (LPPOINT)&pt, *pdwEffect);
		return S_OK;
	}
	STDMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
	{
		if (m_pDropTargetHelper != NULL)
			m_pDropTargetHelper->DragOver((LPPOINT)&pt, *pdwEffect);
		return S_OK;
	}
	STDMETHODIMP DragLeave()
	{
		if (m_pDropTargetHelper != NULL)
			m_pDropTargetHelper->DragLeave();
		return S_OK;
	}
	STDMETHODIMP Drop(IDataObject *pDataObj, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect)
	{
		if (m_pDropTargetHelper != NULL)
			m_pDropTargetHelper->Drop(pDataObj, (LPPOINT)&pt, *pdwEffect);
		return S_OK;
	}
private:
	LONG              m_cRef;
	HWND              m_hwnd;
	IDropTargetHelper *m_pDropTargetHelper;
};

class CDropSource : public IDropSource
{
public:
	CDropSource() { m_cRef = 1; }
	~CDropSource() {}
	STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject)
	{
		*ppvObject = NULL;
		if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDropSource))
			*ppvObject = static_cast<IDropSource *>(this);
		else
			return E_NOINTERFACE;
		AddRef();
		return S_OK;
	}
	STDMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&m_cRef);
	}
	STDMETHODIMP_(ULONG) Release()
	{
		if (InterlockedDecrement(&m_cRef) == 0)
		{
			delete this;
			return 0;
		}
		return m_cRef;
	}
	STDMETHODIMP QueryContinueDrag(BOOL fEscapePressed, DWORD grfKeyState)
	{
		if (fEscapePressed)
			return DRAGDROP_S_CANCEL;
		if ((grfKeyState & MK_LBUTTON) == 0)
			return DRAGDROP_S_DROP;
		return S_OK;
	}
	STDMETHODIMP GiveFeedback(DWORD dwEffect)
	{
		return DRAGDROP_S_USEDEFAULTCURSORS;
	}
private:
	LONG m_cRef;
};

LRESULT CALLBACK ListProc1(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static int m_nMaxRowCount;  // 列数

	static const int m_nItemWidth = DRAG_IMAGE_WIDTH; // 文字列幅+アイコン幅の最大値
	static const int m_nItemHeight = DRAG_IMAGE_HEIGHT; // 項目の高さの最大値（アイコンの高さ含む）

	static const int m_nMinItemOffsetX = 10;
	static int m_nItemOffsetX = m_nMinItemOffsetX;
	static const int m_nItemOutOffsetX = 18;
	static const int m_nMinItemOffsetY = 10;
	static const int m_nItemOutOffsetY = 18;

	static SIZE m_nImageSize;
	static int m_nScrollYPos;        // スクロールY位置
	static int m_nLineCount;          // 全体の行数
	static int m_nAccumDelta;     // for mouse wheel logic

	static COLORREF m_colBackGround = RGB(255, 255, 255);
	static SCROLLINFO si;

	static IDropTarget*pDropTarget;
	static IDropSource*pDropSource;
	static IDragSourceHelper*pDragSourceHelper;

	switch (msg)
	{
	case WM_CREATED:
		{
			pDropTarget = new CDropTarget(hWnd);
			pDropSource = new CDropSource();
			CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pDragSourceHelper));
			RegisterDragDrop(hWnd, pDropTarget);
		}
		break;
	case WM_INVALIDATEITEM:
		{
			const int nIndex = (int)wParam;
			RECT rect;
			rect.left = (nIndex % m_nMaxRowCount) * (m_nItemWidth + m_nItemOffsetX) + m_nItemOutOffsetX;
			rect.right = rect.left + m_nItemWidth + m_nItemOffsetX;
			rect.top = (nIndex / m_nMaxRowCount - m_nScrollYPos) * (m_nItemHeight + m_nMinItemOffsetY) + m_nItemOutOffsetY;
			rect.bottom = rect.top + m_nItemHeight + m_nMinItemOffsetY;
			InvalidateRect(hWnd, &rect, 0);
		}
		break;
	case WM_ENSUREVISIBLE:
		{
			const int nIndex = (int)wParam;
			si.fMask = SIF_POS;
			GetScrollInfo(hWnd, SB_VERT, &si);
			m_nScrollYPos = si.nPos;
			RECT rect;
			GetClientRect(hWnd, &rect);
			if (nIndex / m_nMaxRowCount < m_nScrollYPos)
			{
				si.nPos = nIndex / m_nMaxRowCount;
			}
			else if (nIndex / m_nMaxRowCount - rect.bottom / m_nItemHeight >= m_nScrollYPos)
			{
				si.nPos = nIndex / m_nMaxRowCount - rect.bottom / m_nItemHeight + 1;
			}
			else
			{
				break;
			}
			SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
			GetScrollInfo(hWnd, SB_VERT, &si);
			if (si.nPos != m_nScrollYPos)
			{
				RECT rect;
				GetClientRect(hWnd, &rect);
				rect.top += m_nItemOutOffsetY;
				ScrollWindow(hWnd, 0, m_nItemHeight * (m_nScrollYPos - si.nPos), 0, &rect);
			}
		}
		break;
	case LB_ADDSTRING:
		{
			const int nIndex = (int)CallWindowProc(ListWndProc, hWnd, msg, wParam, lParam);
			SendMessage(hWnd, WM_CALCPARAMETER, 0, 0);
			InvalidateRect(hWnd, 0, 0);
			return nIndex;
		}
		break;
	case WM_MOUSEWHEEL:
		{
			const int nDeltaPerLine = 60;
			m_nAccumDelta += GET_WHEEL_DELTA_WPARAM(wParam);
			while (m_nAccumDelta >= nDeltaPerLine)
			{
				SendMessage(hWnd, WM_VSCROLL, SB_LINEUP, 0);
				m_nAccumDelta -= nDeltaPerLine;
			}
			while (m_nAccumDelta <= -nDeltaPerLine)
			{
				SendMessage(hWnd, WM_VSCROLL, SB_LINEDOWN, 0);
				m_nAccumDelta += nDeltaPerLine;
			}
		}
		break;
	case WM_DESTROY:
		{
			const int nItemCount = (int)SendMessage(hWnd, LB_GETCOUNT, 0, 0);
			for (int nIndex = 0; nIndex < nItemCount; nIndex++)
			{
				LIST_ITEM*itemdata = (LIST_ITEM*)SendMessage(hWnd, LB_GETITEMDATA, nIndex, 0);
				if (itemdata)
				{
					delete itemdata->image;
					delete itemdata;
				}
			}
			if (pDropTarget != NULL) pDropTarget->Release();
			if (pDropSource != NULL) pDropSource->Release();
			if (pDragSourceHelper != NULL) pDragSourceHelper->Release();
			RevokeDragDrop(hWnd);
	}
		break;
	case WM_LBUTTONDOWN:
		{
			const int nItemCount = (int)SendMessage(hWnd, LB_GETCOUNT, 0, 0);
			if (nItemCount == 0)
			{
				return 0;
			}
			const int xPos = lParam & 0xFFFF;
			const int yPos = (lParam >> 16) & 0xFFFF;

			const int xIndex = (xPos - m_nItemOutOffsetX) / (m_nItemWidth + m_nItemOffsetX) + m_nScrollYPos * m_nMaxRowCount;
			const int yIndex = (yPos - m_nItemOutOffsetY) / (m_nItemHeight + m_nMinItemOffsetY);

			const int nOldIndex = (int)SendMessage(hWnd, LB_GETCURSEL, 0, 0);
			const int nIndex = yIndex * m_nMaxRowCount + xIndex;

			if (nOldIndex != nIndex)
			{
				SendMessage(hWnd, LB_SETCURSEL, nIndex, 0);
				if (nOldIndex != LB_ERR)
				{
					SendMessage(hWnd, WM_INVALIDATEITEM, nOldIndex, 0);
				}
				if (nIndex != LB_ERR)
				{
					SendMessage(hWnd, WM_INVALIDATEITEM, nIndex, 0);
					SendMessage(hWnd, WM_ENSUREVISIBLE, nIndex, 0);
				}
			}

			LIST_ITEM*itemdata = (LIST_ITEM*)SendMessage(hWnd, LB_GETITEMDATA, nIndex, 0);
			if (itemdata)
			{
				PIDLIST_ABSOLUTE*ppidlAbsolute = (PIDLIST_ABSOLUTE *)CoTaskMemAlloc(sizeof(PIDLIST_ABSOLUTE));
				PITEMID_CHILD*ppidlChild = (PITEMID_CHILD *)CoTaskMemAlloc(sizeof(PITEMID_CHILD));
				ppidlAbsolute[0] = ILCreateFromPath(itemdata->szImageFilePath);
				IShellFolder*pShellFolder = NULL;
				SHBindToParent(ppidlAbsolute[0], IID_PPV_ARGS(&pShellFolder), NULL);
				ppidlChild[0] = ILFindLastID(ppidlAbsolute[0]);
				IDataObject*pDataObject;
				pShellFolder->GetUIObjectOf(NULL, 1, (LPCITEMIDLIST *)ppidlChild, IID_IDataObject, NULL, (void **)&pDataObject);

				POINT point = { (LONG)(DRAG_IMAGE_WIDTH / 2), (LONG)(DRAG_IMAGE_HEIGHT / 2) };
				SetDragImage(hWnd, pDragSourceHelper, pDataObject, itemdata->image, &point);
				DWORD dwEffect;
				DoDragDrop(pDataObject, pDropSource, DROPEFFECT_COPY, &dwEffect);
				CoTaskMemFree(ppidlAbsolute[0]);
				CoTaskMemFree(ppidlAbsolute);
				CoTaskMemFree(ppidlChild);
			}
		}
		break;
	case WM_SIZE:
		SendMessage(hWnd, WM_CALCPARAMETER, 0, 0);
		InvalidateRect(hWnd, 0, 0);
		break;
	case WM_VSCROLL:
		{
			si.fMask = SIF_ALL;
			GetScrollInfo(hWnd, SB_VERT, &si);
			m_nScrollYPos = si.nPos;
			switch (LOWORD(wParam))
			{
			case SB_TOP:
				si.nPos = si.nMin;
				break;
			case SB_BOTTOM:
				si.nPos = si.nMax;
				break;
			case SB_LINEUP:
				si.nPos -= 1;
				break;
			case SB_LINEDOWN:
				si.nPos += 1;
				break;
			case SB_PAGEUP:
				si.nPos -= si.nPage;
				break;
			case SB_PAGEDOWN:
				si.nPos += si.nPage;
				break;
			case SB_THUMBTRACK:
				si.nPos = HIWORD(wParam);//si.nTrackPos;
				break;
			default:
				break;
			}
			si.fMask = SIF_POS;
			SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
			GetScrollInfo(hWnd, SB_VERT, &si);
			if (si.nPos != m_nScrollYPos)
			{
				RECT rect;
				GetClientRect(hWnd, &rect);
				rect.top += m_nItemOutOffsetY;
				ScrollWindow(hWnd, 0, (m_nItemHeight + m_nMinItemOffsetY) * (m_nScrollYPos - si.nPos), 0, &rect);
			}
		}
		break;
	case WM_ERASEBKGND:
		return TRUE;
	case WM_PAINT:
		{
			PAINTSTRUCT ps;
			HDC hdc = BeginPaint(hWnd, &ps);
			const COLORREF clrPrev = SetBkColor(hdc, m_colBackGround);
			ExtTextOut(hdc, 0, 0, ETO_OPAQUE, &ps.rcPaint, 0, 0, 0);
			SetBkColor(hdc, clrPrev);
			const int nItemCount = (int)SendMessage(hWnd, LB_GETCOUNT, 0, 0);
			if (nItemCount > 0)
			{
				si.fMask = SIF_POS;
				GetScrollInfo(hWnd, SB_VERT, &si);
				m_nScrollYPos = si.nPos;
				const int nStartPos = m_nScrollYPos * m_nMaxRowCount;
				const int nEndPos = nItemCount;
				int nWidth = m_nItemOutOffsetX;
				int nHeight = m_nItemOutOffsetY;
				int nRowCount = 0;
				const int nCursel = (int)SendMessage(hWnd, LB_GETCURSEL, 0, 0);
				for (int i = nStartPos; i < nEndPos; ++i)
				{
					if (nRowCount >= m_nMaxRowCount)
					{
						nRowCount = 0;
						nWidth = m_nItemOutOffsetX;
						nHeight += m_nItemHeight + m_nMinItemOffsetY;
					}
					if (i == nCursel)
					{
						Rectangle(hdc, nWidth, nHeight, nWidth + m_nItemWidth, nHeight + m_nItemHeight);
						//選択されている。
					}
					LIST_ITEM*itemdata = (LIST_ITEM*)SendMessage(hWnd, LB_GETITEMDATA, i, 0);
					if (itemdata)
					{
						Gdiplus::Graphics g(hdc);
						Gdiplus::Rect rect = { nWidth + 3,nHeight + 3, m_nItemWidth - 6, m_nItemHeight - 6 };
						g.DrawImage(itemdata->image, rect);
					}
					++nRowCount;
					nWidth += m_nItemWidth + m_nItemOffsetX;
				}
			}
			EndPaint(hWnd, &ps);
		}
		break;
	case WM_CALCPARAMETER:
		{
			RECT rect;
			GetClientRect(hWnd, &rect);
			const int nItemCount = (int)SendMessage(hWnd, LB_GETCOUNT, 0, 0);
			if (nItemCount == 0)
			{
				m_nMaxRowCount = 1; // 文字列を挿入するかまたはウィンドウサイズを変更するたびに変わる内部変数
				m_nScrollYPos = 0; // スローるバーの位置
				m_nLineCount = 1; // 全体の行数
			}
			else
			{
				// アイテムの最大幅、最大高さを計算
				// 列数を計算
				m_nMaxRowCount = (rect.right - m_nItemOutOffsetX * 2 + m_nMinItemOffsetX) / (m_nItemWidth + m_nMinItemOffsetX);
				if (m_nMaxRowCount == 0)
				{
					m_nMaxRowCount = 1;
					m_nItemOffsetX = m_nMinItemOffsetX;
				}
				else if (m_nMaxRowCount == 1)
				{
					m_nItemOffsetX = m_nMinItemOffsetX;
				}
				else
				{
					m_nItemOffsetX = (rect.right - m_nItemOutOffsetX * 2 - m_nMaxRowCount*m_nItemWidth) / (m_nMaxRowCount - 1);
				}
				// 行数を計算
				m_nLineCount = (nItemCount - 1) / m_nMaxRowCount + 1;
			}
			// スクロールバーを更新
			si.fMask = SIF_RANGE | SIF_PAGE;
			si.nMin = 0;
			si.nMax = m_nLineCount - 1;
			si.nPage = (rect.bottom - m_nItemOutOffsetY * 2 + m_nMinItemOffsetX) / (m_nItemHeight + m_nMinItemOffsetX); // (rect.bottom-m_nItemOutOffsetY*2+m_nMinItemOffsetY) / (m_nItemHeight+m_nMinItemOffsetY);
			SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
		}
		break;
	default:
		return CallWindowProc(ListWndProc, hWnd, msg, wParam, lParam);
	}
	return 0;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hList;
	switch (msg)
	{
	case WM_CREATE:
		OleInitialize(NULL);
		hList = CreateWindow(TEXT("LISTBOX"), 0, WS_VISIBLE | WS_CHILD | LBS_NOREDRAW | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SetClassLongPtr(hList, GCL_STYLE, GetClassLongPtr(hList, GCL_STYLE)&~CS_DBLCLKS);
		ListWndProc = (WNDPROC)SetWindowLongPtr(hList, GWLP_WNDPROC, (LONG_PTR)ListProc1);
		SendMessage(hList, WM_CREATED, 0, 0);
		{
			TCHAR szFindDirectory[MAX_PATH];
			TCHAR szSearchPath[MAX_PATH];
			TCHAR szImagePath[MAX_PATH];
			WIN32_FIND_DATA FindFileData;
			HANDLE hFind = INVALID_HANDLE_VALUE;
			ExpandEnvironmentStrings(TEXT("%localappdata%\\Packages\\Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy\\LocalState\\Assets\\"), szFindDirectory, MAX_PATH);
			lstrcpy(szSearchPath, szFindDirectory);
			lstrcat(szSearchPath, TEXT("*"));
			if (!((hFind = FindFirstFile(szSearchPath, &FindFileData)) == INVALID_HANDLE_VALUE))
			{
				do {
					if (!lstrcmp(FindFileData.cFileName, TEXT(".")) || !lstrcmp(FindFileData.cFileName, TEXT("..")))continue;
					if ( (FindFileData.dwFileAttributes&FILE_ATTRIBUTE_DIRECTORY) == 0)
					{
						lstrcpy(szImagePath, szFindDirectory);
						lstrcat(szImagePath, FindFileData.cFileName);
						Gdiplus::Image *image = new Gdiplus::Image(szImagePath);
						if (image && image->GetLastStatus() == Gdiplus::Status::Ok && image->GetWidth() >= 1920 && image->GetHeight() >= 1080)
						{
							const int nIndex = (int)SendMessage(hList, LB_ADDSTRING, 0, (LPARAM)TEXT(""));
							LIST_ITEM* item = new LIST_ITEM;
							item->image = image->GetThumbnailImage(64, 64);
							lstrcpy(item->szImageFilePath, szImagePath);
							SendMessage(hList, LB_SETITEMDATA, nIndex, (LPARAM)item);
							delete image;
						}
					}
				} while (FindNextFile(hFind, &FindFileData));
				FindClose(hFind);
			}
		}
		break;
	case WM_SETFOCUS:
		SetFocus(hList);
		break;
	case WM_SIZE:
		MoveWindow(hList, 0, 0, LOWORD(lParam), HIWORD(lParam), TRUE);
		break;
	case WM_DESTROY:
		DestroyWindow(hList);
		OleUninitialize();
		PostQuitMessage(0);
		break;
	default:
		return DefWindowProc(hWnd, msg, wParam, lParam);
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	ULONG_PTR gdiToken;
	GdiplusStartupInput gdiSI;
	GdiplusStartup(&gdiToken, &gdiSI, NULL);
	MSG msg;
	WNDCLASS wndclass = {
		0,
		WndProc,
		0,
		0,
		hInstance,
		0,
		LoadCursor(0,IDC_ARROW),
		0,
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("Spotlight Wallpaper"),
		WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	GdiplusShutdown(gdiToken);
	return (int)msg.wParam;
}
