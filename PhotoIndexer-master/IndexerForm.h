#pragma once
//
// IndexerForm.h
//


#include "resource.h"
#include "atlstr.h"
#include <string>
#include <gdiplus.h>
#include "ImageTiler.h"
#include "Thread.h"
#include "stdstringfcn.h"
#include "Utils.h"

using namespace Gdiplus;

class CIndexerForm :
	public CAxDialogImpl<CIndexerForm>
{
public:
	enum { IDD = IDD_DIALOG1 };
	CString sSelectedDir,sDestinationDir;
	CString sPrepend;

	CIndexerForm(void);
	~CIndexerForm(void);
	CEdit edit1;
	CEdit edit2;
	CEdit edit3;
	CButton check1;
	CButton check2;
	CProgressBarCtrl  progress1;
	CButton cancelButton;

	CImageTiler tiler;
	Gdiplus::Bitmap* masterBitmap;
	Gdiplus::Bitmap* masterBitmapSized;

	
	BOOL PreTranslateMessage(MSG* pMsg)
	{
		return CWindow::IsDialogMessage(pMsg);
	}	
	BEGIN_MSG_MAP(CIndexerForm)
		//MESSAGE_HANDLER(WM_CREATE, OnCreateDialog)
		MESSAGE_HANDLER(WM_INITDIALOG, OnCreateDialog)
		
		COMMAND_ID_HANDLER(IDC_DIRBUTTON1, OnFromDirectory)
		COMMAND_ID_HANDLER(IDC_DIRBUTTON2, OnToDirectory)
		COMMAND_ID_HANDLER(IDC_GENERATE, OnCreateIndex)
		COMMAND_ID_HANDLER(IDC_GENERATE2, OnRecusiveIndex)
		COMMAND_ID_HANDLER(IDCANCEL, OnCancelIndexing)
		//COMMAND_ID_HANDLER(IDC_CHECK1, BN_CLICKED, OnCheck1)

		REFLECT_NOTIFICATIONS()
	END_MSG_MAP()
	LRESULT OnCreateDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/);
	LRESULT OnFromDirectory(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnToDirectory(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCreateIndex(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnRecusiveIndex(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);
	LRESULT OnCancelIndexing(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/);

	LRESULT GenerateIndex(std::string selectedFolder, std::string destFolder);
	//DWORD CIndexerForm::RecusiveIndex();
	static void Run(void * lpVoid);
	bool bRunning;
	uintptr_t _threadid;

};
