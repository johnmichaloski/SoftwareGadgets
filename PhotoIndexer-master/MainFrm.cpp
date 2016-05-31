// MainFrm.cpp : implmentation of the CMainFrame class
//
/////////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "resource.h"

#include "aboutdlg.h"
#include "PhotoIndexerView.h"
#include "MainFrm.h"
#include "ImageTiler.h"
#include "Utils.h"
#include <comdef.h>

CMainFrame * mainFrm;


BOOL CMainFrame::PreTranslateMessage(MSG* pMsg)
{
	if(CFrameWindowImpl<CMainFrame>::PreTranslateMessage(pMsg))
		return TRUE;

	return m_view.PreTranslateMessage(pMsg);
}

BOOL CMainFrame::OnIdle()
{
	CWtlHtmlView * view=NULL;

	UIUpdateToolBar();

	int n = m_view.GetActivePage();
	if( n<0 && n >= _pages.size()) 
	{
		UIEnable ( ID_BACK, false );
		UIEnable ( ID_FWD, false );
		return 0;
	}
	if(n>=1)  // zero web browswers at first
		view = _pages[n-1];
	else if(view == NULL)
		n=0;
	if(n<1)
	{
		UIEnable ( ID_BACK, false );
		UIEnable ( ID_FWD, false );	}
	else
	{	
		bool bFwdFlag  = view->CanGoForward();
		UIEnable ( ID_FWD, bFwdFlag );
		bool bBackFlag  = view->CanGoBack();
		UIEnable ( ID_BACK, bBackFlag );
	}
	return FALSE;
}

LRESULT CMainFrame::OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	mainFrm = this;
	// create command bar window
	HWND hWndCmdBar = m_CmdBar.Create(m_hWnd, rcDefault, NULL, ATL_SIMPLE_CMDBAR_PANE_STYLE);
	// attach menu
	m_CmdBar.AttachMenu(GetMenu());
	// load command bar images
	m_CmdBar.LoadImages(IDR_MAINFRAME);
	// remove old menu
	SetMenu(NULL);

	HWND hWndToolBar = CreateSimpleToolBarCtrl(m_hWnd, IDR_MAINFRAME, FALSE, ATL_SIMPLE_TOOLBAR_PANE_STYLE);

	CreateSimpleReBar(ATL_SIMPLE_REBAR_NOBORDER_STYLE);
	AddSimpleReBarBand(hWndCmdBar);
	AddSimpleReBarBand(hWndToolBar, NULL, TRUE);

	CreateSimpleStatusBar();

	m_hWndClient = m_view.Create(m_hWnd, rcDefault, NULL, WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | WS_CLIPCHILDREN, WS_EX_CLIENTEDGE);

	UIAddToolBar(hWndToolBar);
	UISetCheck(ID_VIEW_TOOLBAR, 1);
	UISetCheck(ID_VIEW_STATUS_BAR, 1);

	// register object for message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->AddMessageFilter(this);
	pLoop->AddIdleHandler(this);

	m_photoview = new CIndexerForm ;
	m_photoview->ATL::CAxDialogImpl<CIndexerForm,ATL::CWindow>::Create(m_view, rcDefault);
	m_view.AddPage(m_photoview->m_hWnd, _T("Home"));

	this->SetWindowText("Photo Indexer");

    UIEnable ( ID_BACK, false );
    UIEnable ( ID_FWD, false );
 
	return 0;
}

LRESULT CMainFrame::OnDestroy(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& bHandled)
{
	// unregister message filtering and idle updates
	CMessageLoop* pLoop = _Module.GetMessageLoop();
	ATLASSERT(pLoop != NULL);
	pLoop->RemoveMessageFilter(this);
	pLoop->RemoveIdleHandler(this);

	bHandled = FALSE;
	return 1;
}
LRESULT CMainFrame::OnFilePrint(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	unsigned int n = m_view.GetActivePage();
	if( n<0 && n > _pages.size()) return 0;
	CWtlHtmlView * pHtmlView = _pages[n-1];
	_variant_t vArg;
	pHtmlView->ExecCommand(OLECMDID_PRINT,OLECMDEXECOPT_DODEFAULT, NULL, &vArg); 

	return 0;
}
LRESULT CMainFrame::OnFilePrintPreview(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	unsigned int n = m_view.GetActivePage();
	if( n<1 && n >= _pages.size()) return 0;
	CWtlHtmlView * pHtmlView = _pages[n-1];
	_variant_t vArg;
	pHtmlView->ExecCommand(OLECMDID_PRINTPREVIEW,OLECMDEXECOPT_DODEFAULT, NULL, &vArg); 

	return 0;
}
LRESULT CMainFrame::OnFileExit(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	PostMessage(WM_CLOSE);
	return 0;
}

LRESULT CMainFrame::OnFileNew(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	// TODO: add code to initialize document

	return 0;
}

LRESULT CMainFrame::OnViewToolBar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	static BOOL bVisible = TRUE;	// initially visible
	bVisible = !bVisible;
	CReBarCtrl rebar = m_hWndToolBar;
	int nBandIndex = rebar.IdToIndex(ATL_IDW_BAND_FIRST + 1);	// toolbar is 2nd added band
	rebar.ShowBand(nBandIndex, bVisible);
	UISetCheck(ID_VIEW_TOOLBAR, bVisible);
	UpdateLayout();
	return 0;
}

LRESULT CMainFrame::OnViewStatusBar(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	BOOL bVisible = !::IsWindowVisible(m_hWndStatusBar);
	::ShowWindow(m_hWndStatusBar, bVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
	UISetCheck(ID_VIEW_STATUS_BAR, bVisible);
	UpdateLayout();
	return 0;
}

LRESULT CMainFrame::OnAppAbout(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CAboutDlg dlg;
	dlg.DoModal();
	return 0;
}


LRESULT CMainFrame::OnShowIndex(UINT /*uMsg*/, WPARAM wParam, LPARAM lParam, BOOL& /*bHandled*/)
{
	CWtlHtmlView * pHtmlView = new CWtlHtmlView();
	std::string path = (LPCSTR) lParam;  // lparam - path
	std::string  title = (LPCSTR) wParam;  // wparam - title
	pHtmlView->Create(m_view,rcDefault,
		path.c_str(), 
		WS_CHILD | WS_VISIBLE | WS_VSCROLL,
		WS_EX_CLIENTEDGE);

	m_view.AddPage(pHtmlView->m_hWnd,title.c_str());
	_pages.push_back(pHtmlView);
	delete (LPCSTR) lParam;
	delete (LPCSTR) wParam;
	return 0;
}

LRESULT CMainFrame::OnShowStatus(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*bHandled*/)
{
	char * msg = (char *) lParam;
	CWindow(m_hWndStatusBar).SetWindowText(msg);
	return 0;
}

LRESULT CMainFrame::OnGeneratePhotos(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*bHandled*/)
{
	std::string tmp;
	int numcols=6;
	tmp+="<HTML><BODY><TABLE>";
	//m_photoview
	std::string destFolder = (LPCSTR) m_photoview->sDestinationDir;
	std::vector<std::string> files= Utils::FileList(destFolder);
	for(unsigned int i=0; i< files.size(); i=i+numcols)
	{
		tmp+="<TR>";
		for(int j=0; j< numcols; j++)
		{
			if((i+j) >= files.size())
				break;
			tmp+=StdStringFormat("<TD><A HREF=\"%s\"><IMG SRC=\"%s\" width=\"%d\" height=\"%d\" ><BR>%s</A></TD>\n", 
				files[i+j].c_str(), files[i+j].c_str(), 200, 200, files[i+j].c_str());
		}	
		tmp+="</TR>";

	}
	tmp+="</TABLE></BODY></HTML>";

	WriteFile(::ExeDirectory() + "indexingtmp.html", tmp);
	m_wtlview = new CWtlHtmlView ;
	m_wtlview->Create(m_view, rcDefault,(::ExeDirectory() +"indexingtmp.html").c_str(),	WS_CHILD | WS_VISIBLE | WS_VSCROLL,
		WS_EX_CLIENTEDGE);
	m_view.AddPage(m_wtlview->m_hWnd, ExtractFilename(destFolder).c_str());
	_pages.push_back (m_wtlview);

	m_wtlview->CanGoForward();

	//m_wtlview->SetDocumentText(tmp.c_str());
	return 0;
}

LRESULT CMainFrame::OnGoBack(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{

	m_wtlview->CanGoForward();
	return 0;
}