#include "StdAfx.h"
#include "ImageTiler.h"
#include <comdef.h>
#include "atlstr.h"

#include <atlconv.h>
#pragma comment (lib, "Gdiplus.lib")

CImageTiler::CImageTiler(void)
{
	gpBitmap=NULL;
}

CImageTiler::~CImageTiler(void)
{
}
void CImageTiler::Init()
{	// Initialize GDI+
	GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);

	gpBitmap = NULL;

	
}

void CImageTiler::Shutdown()
{
	if (gpBitmap) delete gpBitmap;

	GdiplusShutdown(gdiplusToken);
}

// Create an Image object based on a PNG file.
Gdiplus::Bitmap * CImageTiler::ReadJpgImg(std::string szJpgFileName) // as bmp
{
	bstr_t file((LPCSTR) szJpgFileName.c_str());
	gpBitmap = new Gdiplus::Bitmap(file);
	return gpBitmap;
}

void CImageTiler::SaveBmpAsJpg(Bitmap *  bitmap,std::string szJpgFileName)
{
	static const char* StatusMsgMap[] = 
{
    "Ok",               //StatusMsgMap[Ok] = "Ok";
    "GenericError",     //StatusMsgMap[GenericError] = "GenericError";
    "InvalidParameter", //StatusMsgMap[InvalidParameter] = "InvalidParameter";
    "OutOfMemory",      //StatusMsgMap[OutOfMemory] = "OutOfMemory";
	"ObjectBusy = 4",
      "InsufficientBuffer = 5",
    "NotImplemented = 6",
    "Win32Error = 7",
    "WrongState = 8",
    "Aborted = 9",
    "FileNotFound = 10",
    "ValueOverflow = 11",
    "AccessDenied = 12",
    "UnknownImageFormat = 13",
    "FontFamilyNotFound = 14",
    "FontStyleNotFound = 15",
   "NotTrueTypeFont = 16",
    "UnsupportedGdiplusVersion = 17",
    "GdiplusNotInitialized = 18",
    "PropertyNotFound = 19",
    "PropertyNotSupported = 20",
};

	//Image  image(bitmap);
	// Save the image.
	CLSID jpgClsid;
	GetEncoderClsid(L"image/jpeg", &jpgClsid);
	Status status = bitmap->Save(bstr_t(szJpgFileName.c_str()), &jpgClsid, NULL);
	if(status != 0)
	{
			if(status< 21)
				OutputDebugString(StatusMsgMap[status]);
	}
}

Gdiplus::Bitmap* CImageTiler::ResizeClone(Bitmap *bmp, INT width, INT height)
{
    UINT o_height = bmp->GetHeight();
    UINT o_width = bmp->GetWidth();
    INT n_width = width;
    INT n_height = height;
    double ratio = ((double)o_width) / ((double)o_height);
    if (o_width > o_height) {
        // Resize down by width
        n_height = static_cast<UINT>(((double)n_width) / ratio);
    } else {
        n_width = static_cast<UINT>(n_height * ratio);
    }
    Gdiplus::Bitmap* newBitmap = new Gdiplus::Bitmap(n_width, n_height, bmp->GetPixelFormat());
    Gdiplus::Graphics graphics(newBitmap);
    graphics.DrawImage(bmp, 0, 0, n_width, n_height);
    return newBitmap;
}
// assumes photots in via bitmap, photos out...
Gdiplus::Bitmap* CImageTiler::CreateCompatibleBitmap(Bitmap * bitmap)
{
	HBITMAP  memDC ;
	Gdiplus::Color clr(0xFF,0xFF,0xFF); // (only blue component gets used)
	bitmap->GetHBITMAP(clr,&memDC);
	long width = bitmap->GetWidth();
	long height = bitmap->GetHeight();
	Gdiplus::PixelFormat pixelFormat = bitmap->GetPixelFormat();

    Gdiplus::Bitmap* newBitmap = new Gdiplus::Bitmap(width, height, pixelFormat);
    Gdiplus::Graphics graphics(newBitmap);
 //   graphics.DrawImage(bitmap, 0, 0, width, height);
	graphics.Clear(Color::White);
    return newBitmap;
}

void CImageTiler::AddTitle(Graphics * graphics, std::string title)
{
	bstr_t text(title.c_str());

	SolidBrush whiteBrush(Color(255, 255, 255, 255));

	Gdiplus::Font myFont(L"Arial", 80,FontStyleBold);
	PointF origin(150.0f, 10.0f);
	SolidBrush blackBrush(Color(255, 0, 0, 0));
	graphics->DrawString(text,wcslen(text), &myFont, origin, &whiteBrush);
}
Gdiplus::Bitmap* CImageTiler::TileImage(
	Bitmap * DestBmp, 
	std::vector<Bitmap  *> SrcBmps,
	std::string title)
{
	if (DestBmp == NULL)
		throw std::exception("CImageTiler::TileImage No destination bitmap");

	if(SrcBmps.size() < 1)
		throw std::exception("CImageTiler::TileImage SrcBmps no elements");

	Graphics * graphics = new Graphics(DestBmp);
	
	if (graphics == NULL)
		throw std::exception("CImageTiler::TileImage No graphics from bitmap");

	graphics->Clear(Color::White);

	int THUMBNAILX= (DestBmp->GetWidth() - (9 * 5))/6; // 80;
	int THUMBNAILY= (DestBmp->GetHeight() - (7 * 5))/6;  //60;
	int SPACEX= 5;
	int SPACEY= 5;
	int x=SPACEX*2;
	int y=SPACEY*2;
	int TILESX=6;
	int TILESY=6;
	for(int k=0; k< SrcBmps.size(); k++)
	{
		int i=(k%TILESX);
		if((k%TILESX)==0  && k!=0)
			y=y+THUMBNAILY+SPACEY;

		graphics->SetInterpolationMode(InterpolationModeHighQualityBicubic);
		graphics->DrawImage(
			SrcBmps[k],
			(REAL)x+i*(THUMBNAILX+SPACEX),
			(REAL)y,
			(REAL)THUMBNAILX,
			(REAL)THUMBNAILY );
		
		//StretchBlt (hdcDest,
		//x+i*110, y, 100, 140,
		//hdcMem,
		//0, 0,SrcBmps[k].GetWidth(), SrcBmps[k].GetHeight(), //  bm.bmWidth, bm.bmHeight
		//SRCCOPY);


	}
	if(!title.empty())
	{
		AddTitle(graphics, title);
	}
	return DestBmp;
}
//bool CImageTiler::TileImage(
//	Bitmap &SrcBmp, 
//	const std::wstring& szDestFile,
//	const std::wstring& szEncoderString )
//{
//	Bitmap DestBmp(
//		SrcBmp.GetWidth()*3,
//		SrcBmp.GetHeight()*3,
//		PixelFormat32bppARGB );
//
//	Graphics graphics(&DestBmp);
//
//	for( int x=0; x<3; ++x )
//	{
//		for( int y=0; y<3; ++y )
//		{
//			graphics.DrawImage(
//				&SrcBmp,
//				(REAL)x*SrcBmp.GetWidth(),
//				(REAL)y*SrcBmp.GetHeight(),
//				(REAL)SrcBmp.GetWidth(),
//				(REAL)SrcBmp.GetHeight() );
//		}
//	}
//
//// After you got the tiling DestBmp, do what you want, eg displaying or saving to a file
//
//	CLSID Clsid;
//	int result = GetEncoderClsid(szEncoderString.c_str(), &Clsid);
//
//	if( result < 0 )
//		return false;
//
//	Status status = DestBmp.Save( szDestFile.c_str(), &Clsid );
//
//	return status == Ok;
//}

int CImageTiler::GetEncoderClsid(const WCHAR* format, CLSID* pClsid)
{
   UINT  num = 0;          // number of image encoders
   UINT  size = 0;         // size of the image encoder array in bytes

   ImageCodecInfo* pImageCodecInfo = NULL;

   GetImageEncodersSize(&num, &size);
   if(size == 0)
      return -1;  // Failure

   pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
   if(pImageCodecInfo == NULL)
      return -1;  // Failure

   GetImageEncoders(num, size, pImageCodecInfo);

   for(UINT j = 0; j < num; ++j)
   {
      if( wcscmp(pImageCodecInfo[j].MimeType, format) == 0 )
      {
         *pClsid = pImageCodecInfo[j].Clsid;
         free(pImageCodecInfo);
         return j;  // Success
      }    
   }

   free(pImageCodecInfo);
   return -1;  // Failure
}