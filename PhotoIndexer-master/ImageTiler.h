#pragma once

#include <string>
#include <vector>

#include <Gdiplus.h>
using namespace Gdiplus;


/**
Bitmap SrcBmp(L"E:\\hh\\GreenRect.bmp", TRUE);

	bool bRet = TileImage(
		SrcBmp, 
		L"E:\\GreenRect.bmp",
		L"image/bmp" );

	if( bRet )
		MessageBox(NULL, L"Successful", L"Title", MB_OK);
	else
		MessageBox(NULL, L"Failed", L"Title", MB_OK);
*/
class CImageTiler
{
public:
	CImageTiler(void);
	~CImageTiler(void);
	//bool TileImage(
	//	Bitmap &SrcBmp, 
	//	const std::wstring& szDestFile,
	//	const std::wstring& szEncoderString );
	int GetEncoderClsid(const WCHAR* format, CLSID* pClsid);
	void Init();
	Gdiplus::Bitmap * ReadJpgImg(std::string szJpgFileName) ;
	void SaveBmpAsJpg(Bitmap *  bitmap,std::string szJpgFileName);
	void Shutdown();
	Gdiplus::Bitmap* ResizeClone(Bitmap *bmp, INT width, INT height);
	Gdiplus::Bitmap* CreateCompatibleBitmap(Bitmap * bitmap);
	Gdiplus::Bitmap* TileImage(		
		Bitmap * DestBmp, 
		std::vector<Bitmap  *> SrcBmps,
		std::string title="");

	void AddTitle(Graphics * graphics, std::string title);

	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR           gdiplusToken;
	Bitmap *gpBitmap;								// The bitmap for displaying an image

};
