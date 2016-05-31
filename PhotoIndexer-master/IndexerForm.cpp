//
// IndexerForm.cpp
//

#include "StdAfx.h"
#include "IndexerForm.h"
#include <vector>
#include "MainFrm.h"
#include <comdef.h>

extern CMainFrame * mainFrm;

static std::string JustTitle(std::string path)
{

	std::string::size_type n;
	if((n=path.rfind("\\"))!= std::string::npos)
	{
		return path.substr(n+1);
	}

	return path;
}

CIndexerForm::CIndexerForm(void)
{
	CImageTiler().Init();
	masterBitmapSized=Bitmap::FromFile(bstr_t((::ExeDirectory() + "Happy Easter Mom Dad Smiling.JPG").c_str()));
	masterBitmap = tiler.CreateCompatibleBitmap(masterBitmapSized);  // for now...

	TCHAR mypicturespath[MAX_PATH];
    HRESULT result = SHGetFolderPath(NULL, CSIDL_MYPICTURES, NULL, SHGFP_TYPE_CURRENT, mypicturespath); 
	sSelectedDir=sDestinationDir=	mypicturespath;
}

CIndexerForm::~CIndexerForm(void)
{
	delete masterBitmapSized;
	CImageTiler().Shutdown();
}

LRESULT CIndexerForm::OnCreateDialog(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
{
	edit1=::GetDlgItem(this->m_hWnd,IDC_EDIT1);
	edit2=::GetDlgItem(this->m_hWnd,IDC_EDIT2);
	edit3=::GetDlgItem(this->m_hWnd,IDC_EDIT3);
	edit1.SetWindowText(sSelectedDir);
	edit2.SetWindowText(sDestinationDir);
	edit3.SetWindowText("");
	check1=::GetDlgItem(this->m_hWnd,IDC_CHECK1);
	check2=::GetDlgItem(this->m_hWnd,IDC_CHECK2);
	check1.SetCheck( BST_UNCHECKED);
	check2.SetCheck( BST_UNCHECKED);
	
	progress1=::GetDlgItem(this->m_hWnd,IDC_PROGRESS1);
	cancelButton=::GetDlgItem(this->m_hWnd,IDCANCEL);
	cancelButton.ShowWindow(SW_HIDE);
	progress1.ShowWindow(SW_HIDE);
	return 0;

}
LRESULT CIndexerForm::OnFromDirectory(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	CFolderDialog  fldDlg(
    this->m_hWnd, //hWndParent
    "Select Photo Folder to Index (Recursive)",
    BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE  );
	fldDlg.SetInitialFolder(sSelectedDir);
	if(fldDlg.DoModal() == IDOK)
	{
		sSelectedDir = (LPCSTR) fldDlg.m_szFolderPath;
		edit1.SetWindowTextA(sSelectedDir);
	}
	return 0;
}
LRESULT CIndexerForm::OnToDirectory(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
CFolderDialog  fldDlg(
    this->m_hWnd, //hWndParent
    "Select Destination Photo Folder",
    BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE  );
	fldDlg.SetInitialFolder(sDestinationDir);
	if(fldDlg.DoModal() == IDOK)
	{
		sDestinationDir = (LPCSTR) fldDlg.m_szFolderPath;
		edit2.SetWindowTextA(sDestinationDir);
	}
	return 0;
}
LRESULT CIndexerForm::OnCreateIndex(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	std::string folder = (LPCSTR) sSelectedDir;
	std::string destfolder = (LPCSTR) sDestinationDir;
	GenerateIndex(folder, destfolder);
	return 0;
}
//DWORD WINAPI ThreadFunction(LPVOID);




LRESULT CIndexerForm::GenerateIndex(std::string selectedFolder, std::string destFolder)
{	
	std::vector<Bitmap * > bitmaps;

	//	std::string folder="C:\\Program Files\\NIST\\proj\\MTConnect\\Nist\\MTConnectGadgets\\PhotoIndexer\\Debug\\2010 UMDGrad CememonyII Party";
	//	std::string folder2="C:\\Documents and Settings\\michalos\\My Documents\\My Pictures\\Work\\ISD";
	if(selectedFolder.empty())
		return 0;

	std::vector<std::string> allfiles= Utils::FileList(selectedFolder);

	int k=0;
	std::vector<std::string> files;
	for(UINT i=0; i< allfiles.size(); i+=36)
	{
		int j=MIN(allfiles.size(), i+36);
		files.clear();
		files.insert(files.begin(), allfiles.begin()+i, allfiles.begin()+j);
		//tiler.Init();

		// C:\Documents and Settings\michalos\My Documents\My Pictures\Work\ISD
		Bitmap * bitmap=NULL;
		bitmaps.clear();
		for(UINT  j=0; j< files.size(); j++)
		{
			bitmap = tiler.ReadJpgImg(files[j]) ;
			if(bitmap!=NULL)
				bitmaps.push_back(bitmap);
		}

		if(bitmaps.size() < 1)
			return 0;

		std::string title = JustTitle(selectedFolder);
		if(k>0)
			title+=StdStringFormat("_%d",k);
		k++;
		Bitmap * finalbitmap= tiler.TileImage(masterBitmap, 
			bitmaps,
			title);

		//destfolder+="//" + title;
		CreateDirectory(destFolder.c_str(), NULL);
		tiler.SaveBmpAsJpg(finalbitmap,destFolder + "\\"+ title + ".jpg");
		// Delete files
		for(UINT n=0; n< bitmaps.size(); n++)
			delete bitmaps[n];
	}
	//char * szTitle = new char[title.size() +2];
	//strncpy(szTitle, title.c_str(), title.size()+2);

	//std::string path = destFolder + "\\"+ title + ".jpg";
	//char * szPath = new char[path.size() +2];
	//strncpy(szPath, path.c_str(), path.size()+1);

//	::SendMessage(mainFrm->m_hWnd, WM_APP,(WPARAM) szTitle, (LPARAM)szPath);
	return 0;
}	
LRESULT CIndexerForm::OnCancelIndexing(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	bRunning=false;
	//::TerminateThread(_threadid,0);

	return 0;
}
LRESULT CIndexerForm::OnRecusiveIndex(WORD /*wNotifyCode*/, WORD /*wID*/, HWND /*hWndCtl*/, BOOL& /*bHandled*/)
{
	_beginthread( CIndexerForm::Run, 0, (void *) this );
	return 0;
}
void CIndexerForm::Run(void * lpVoid)
{
	CIndexerForm * indexer= (CIndexerForm *) lpVoid;

	indexer->cancelButton.ShowWindow(SW_SHOWNORMAL);
	indexer->progress1.ShowWindow(SW_SHOWNORMAL);

	std::string selectedFolder = (LPCSTR) indexer->sSelectedDir;
	std::string destFolder = (LPCSTR) indexer->sDestinationDir;
	std::vector<std::string> folders= Utils::DirList(selectedFolder);

	indexer->bRunning=true;
	indexer->progress1.ShowWindow(WM_SHOWWINDOW);
	indexer->progress1.SetRange32(0, folders.size());
	for(int i=0; i< folders.size() && indexer->bRunning; i++)
	{
		try{
			indexer->progress1.SetPos(i+1);
			OutputDebugString(StdStringFormat("GenerateIndex for %s  # %d\n", 
				folders[i].c_str(), i).c_str());
			::SendMessage(mainFrm->m_hWnd, SHOW_STATUS_MESSAGE,NULL, (LPARAM) folders[i].c_str());
			indexer->GenerateIndex(folders[i], destFolder);
		}
		catch(...)
		{
			OutputDebugString(StdStringFormat("GenerateIndex exception in %s \n", folders[i].c_str()).c_str());
			DebugBreak();
		}
	}
	indexer->cancelButton.ShowWindow(SW_HIDE);
	indexer->progress1.ShowWindow(SW_HIDE);
	::SendMessage(mainFrm->m_hWnd, SHOW_STATUS_MESSAGE,NULL, (LPARAM) "Ready");
	::SendMessage(mainFrm->m_hWnd, GENERATE_HTML_MESSAGE,NULL, NULL);

}