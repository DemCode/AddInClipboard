
#include "CAddInClipboard.h"

#include <gdiplus.h>

#pragma comment(lib, "gdiplus.lib")

CAddInClipboard::CAddInClipboard(void)
{
	m_iMemory = 0;
	m_iConnect = 0;
}

CAddInClipboard::~CAddInClipboard()
{
}

bool ADDIN_API CAddInClipboard::Init(void* disp)
{
	m_iConnect = (IAddInDefBase*)disp;
	return m_iConnect != NULL;
}

bool ADDIN_API CAddInClipboard::setMemManager(void* memManager)
{
	m_iMemory = (IMemoryManager*)memManager;
	return m_iMemory != 0;
}

long ADDIN_API CAddInClipboard::GetInfo()
{
	return 1000;
}

void ADDIN_API CAddInClipboard::Done()
{
}

bool ADDIN_API CAddInClipboard::RegisterExtensionAs(WCHAR_T** wsExtName)
{
	const wchar_t* wsExtension = L"AddInClipboard";
	size_t iActualSize = ::wcslen(wsExtension) + 1;
	WCHAR_T* dest = 0;

	if (m_iMemory)
	{
		if (m_iMemory->AllocMemory((void**)wsExtName, iActualSize * sizeof(WCHAR_T))) {
			if (!*wsExtName)
				*wsExtName = new WCHAR_T[iActualSize];
			wcscpy(*wsExtName, wsExtension);
		}

		return true;
	}

	return false;
}

long ADDIN_API CAddInClipboard::GetNProps()
{
	return 0;
}

long ADDIN_API CAddInClipboard::FindProp(const WCHAR_T* wsPropName)
{
	return -1;
}

const WCHAR_T* ADDIN_API CAddInClipboard::GetPropName(long lPropNum, long lPropAlias)
{
	return NULL;
}

bool ADDIN_API CAddInClipboard::GetPropVal(const long lPropNum, tVariant* pvarPropVal)
{
	return false;
}

bool ADDIN_API CAddInClipboard::SetPropVal(const long lPropNum, tVariant* pvarPropVal)
{
	return false;
}

bool ADDIN_API CAddInClipboard::IsPropReadable(const long lPropNum)
{
	return false;
}

bool ADDIN_API CAddInClipboard::IsPropWritable(const long lPropNum)
{
	return false;
}

long ADDIN_API CAddInClipboard::GetNMethods()
{
	return 4;
}

long ADDIN_API CAddInClipboard::FindMethod(const WCHAR_T* wsMethodName)
{
	if (wcscmp(wsMethodName, L"ContainsText") == 0 || wcscmp(wsMethodName, L"СодержитТекст") == 0)
		return 0;
	else if (wcscmp(wsMethodName, L"ContainsImage") == 0 || wcscmp(wsMethodName, L"СодержитИзображение") == 0)
		return 1;
	else if (wcscmp(wsMethodName, L"GetText") == 0 || wcscmp(wsMethodName, L"ПолучитьТекст") == 0)
		return 2;
	else if (wcscmp(wsMethodName, L"GetImage") == 0 || wcscmp(wsMethodName, L"ПолучитьИзображение") == 0)
		return 3;

	return -1;
}

const WCHAR_T* ADDIN_API CAddInClipboard::GetMethodName(const long lMethodNum, const long lMethodAlias)
{
	WCHAR_T* result = NULL;

	if (lMethodNum == 0)
		if (lMethodAlias == 0) {
			size_t actualSize = wcslen(L"ContainsText") + 1;
			if (m_iMemory->AllocMemory((void**)&result, actualSize * sizeof(WCHAR_T)))
				wcscpy_s(result, actualSize, L"ContainsText");
		}
		else {
			size_t actualSize = wcslen(L"СодержитТекст") + 1;
			if (m_iMemory->AllocMemory((void**)&result, actualSize * sizeof(WCHAR_T)))
				wcscpy_s(result, actualSize, L"СодержитТекст");
		}
	else if (lMethodNum == 1)
		if (lMethodAlias == 0) {
			size_t actualSize = wcslen(L"ContainsImage") + 1;
			if (m_iMemory->AllocMemory((void**)&result, actualSize * sizeof(WCHAR_T)))
				wcscpy_s(result, actualSize, L"ContainsImage");
		}
		else {
			size_t actualSize = wcslen(L"СодержитИзображение") + 1;
			if (m_iMemory->AllocMemory((void**)&result, actualSize * sizeof(WCHAR_T)))
				wcscpy_s(result, actualSize, L"СодержитИзображение");
		}
	else if (lMethodNum == 2)
		if (lMethodAlias == 0) {
			size_t actualSize = wcslen(L"GetText") + 1;
			if (m_iMemory->AllocMemory((void**)&result, actualSize * sizeof(WCHAR_T)))
				wcscpy_s(result, actualSize, L"GetText");
		}
		else {
			size_t actualSize = wcslen(L"ПолучитьТекст") + 1;
			if (m_iMemory->AllocMemory((void**)&result, actualSize * sizeof(WCHAR_T)))
				wcscpy_s(result, actualSize, L"ПолучитьТекст");
		}
	else if (lMethodNum == 3)
		if (lMethodAlias == 0) {
			size_t actualSize = wcslen(L"GetImage") + 1;
			if (m_iMemory->AllocMemory((void**)&result, actualSize * sizeof(WCHAR_T)))
				wcscpy_s(result, actualSize, L"GetImage");
		}
		else {
			size_t actualSize = wcslen(L"ПолучитьИзображение") + 1;
			if (m_iMemory->AllocMemory((void**)&result, actualSize * sizeof(WCHAR_T)))
				wcscpy_s(result, actualSize, L"ПолучитьИзображение");
		}

	return result;
}

long ADDIN_API CAddInClipboard::GetNParams(const long lMethodNum)
{
	if (lMethodNum == 0)
		return 0l;
	else if (lMethodNum == 1)
		return 0l;
	else if (lMethodNum == 2)
		return 0l;
	else if (lMethodNum == 3)
		return 0l;

	return 0;
}

bool ADDIN_API CAddInClipboard::GetParamDefValue(const long lMethodNum, const long lParamNum, tVariant* pvarParamDefValue)
{
	if (lMethodNum == 0)
		TV_VT(pvarParamDefValue) = VTYPE_EMPTY;
	else if (lMethodNum == 1)
		TV_VT(pvarParamDefValue) = VTYPE_EMPTY;
	else if (lMethodNum == 2)
		TV_VT(pvarParamDefValue) = VTYPE_EMPTY;
	else if (lMethodNum == 3)
		TV_VT(pvarParamDefValue) = VTYPE_EMPTY;

	return true;
}

bool ADDIN_API CAddInClipboard::HasRetVal(const long lMethodNum)
{
	if (lMethodNum == 0)
		return true;
	else if (lMethodNum == 1)
		return true;
	else if (lMethodNum == 2)
		return true;
	else if (lMethodNum == 3)
		return true;

	return false;
}

bool ADDIN_API CAddInClipboard::CallAsProc(const long lMethodNum, tVariant* paParams, const long lSizeArray)
{
	return false;
}

// get image encoder CLSID by its name
int GetEncoderClsid(const wchar_t* format, CLSID* pClsid) {
	unsigned int num = 0, size = 0;
	Gdiplus::GetImageEncodersSize(&num, &size);
	
	if (size == 0) 
		return -1;

	Gdiplus::ImageCodecInfo* pImageCodecInfo = (Gdiplus::ImageCodecInfo*)(malloc(size));
	
	if (pImageCodecInfo == NULL) 
		return -1;
	
	Gdiplus::GetImageEncoders(num, size, pImageCodecInfo);
	
	for (unsigned int j = 0; j < num; ++j) {
		if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0) {
			*pClsid = pImageCodecInfo[j].Clsid;
			free(pImageCodecInfo);
			return j;
		}
	}

	free(pImageCodecInfo);
	return -1;
}

bool ADDIN_API CAddInClipboard::CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{
	if (lMethodNum == 0) {
		TV_VT(pvarRetValue) = VTYPE_BOOL;
		TV_BOOL(pvarRetValue) = containsText();
		return true;
	}
	else if (lMethodNum == 1) {
		TV_VT(pvarRetValue) = VTYPE_BOOL;
		TV_BOOL(pvarRetValue) = containsImage();
		return true;
	}
	else if (lMethodNum == 2) {
		const wchar_t* text = getText();
		WCHAR_T* result = NULL;
		size_t actualSize = wcslen(text) + 1;

		TV_VT(pvarRetValue) = VTYPE_PWSTR;
		if (m_iMemory->AllocMemory((void**) &result, actualSize * sizeof(WCHAR_T)))
		{
			wcscpy_s(result, actualSize, text);
			pvarRetValue->pwstrVal = result;
			pvarRetValue->strLen = wcslen(text);
		}
		return true;
	}
	else if (lMethodNum == 3) {
		TV_VT(pvarRetValue) = VTYPE_EMPTY;
		
		BITMAPINFO* info = getImage();
		
		Gdiplus::Bitmap* bmp = Gdiplus::Bitmap::FromBITMAPINFO(info, (void*)(info + 1));
		
		CLSID imageCLSID;
		GetEncoderClsid(L"image/jpeg", &imageCLSID);
		IStream* pistream = NULL;

		if (CreateStreamOnHGlobal(NULL, true, &pistream) == S_OK) {
			if (bmp->Save(pistream, &imageCLSID) == Gdiplus::Ok) {
				ULARGE_INTEGER ulnSize;
				LARGE_INTEGER lnOffset;
				lnOffset.QuadPart = 0;

				if (pistream->Seek(lnOffset, STREAM_SEEK_END, &ulnSize) == S_OK) {
					if (pistream->Seek(lnOffset, STREAM_SEEK_SET, NULL) == S_OK) {
						char* result = NULL;

						if (m_iMemory->AllocMemory((void**)&result, ulnSize.QuadPart)) {
							ULONG ulBytesRead = 0;

							if (pistream->Read((void*)result, ulnSize.QuadPart, &ulBytesRead) == S_OK) {
								TV_VT(pvarRetValue) = VTYPE_BLOB;
								pvarRetValue->pstrVal = result;
								pvarRetValue->strLen = ulBytesRead;
							}
						}
					}
				}

				pistream->Release();
			}
		}
		return true;
	}

	return false;
}

void ADDIN_API CAddInClipboard::SetLocale(const WCHAR_T* wsLocale)
{
	return void();
}

bool CAddInClipboard::containsText()
{
	return IsClipboardFormatAvailable(CF_TEXT) || IsClipboardFormatAvailable(CF_UNICODETEXT);
}

bool CAddInClipboard::containsImage()
{
	return IsClipboardFormatAvailable(CF_BITMAP) || IsClipboardFormatAvailable(CF_DIB) || IsClipboardFormatAvailable(CF_DIBV5);
}

const wchar_t* CAddInClipboard::getText()
{
	wchar_t* result = NULL;

	if (IsClipboardFormatAvailable(CF_UNICODETEXT))
	{
		if (OpenClipboard(NULL))
		{
			HANDLE hClipboard = GetClipboardData(CF_UNICODETEXT);

			if (hClipboard != NULL && hClipboard != INVALID_HANDLE_VALUE)
			{
				void* dib = GlobalLock(hClipboard);

				if (dib)
				{
					const wchar_t* info = reinterpret_cast<const wchar_t*>(dib);
					result = new wchar_t[wcslen(info) + 1];
					wcscpy(result, info);

					GlobalUnlock(dib);
				}
			}

			CloseClipboard();
		}
	}
	else if (IsClipboardFormatAvailable(CF_TEXT))
	{
		if (OpenClipboard(NULL))
		{
			HANDLE hClipboard = GetClipboardData(CF_TEXT);

			if (hClipboard != NULL && hClipboard != INVALID_HANDLE_VALUE)
			{
				void* dib = GlobalLock(hClipboard);

				if (dib)
				{
					char* info = reinterpret_cast<char*>(dib);

					MultiByteToWideChar(CP_ACP, 0, info, -1, result, 0);

					GlobalUnlock(dib);
				}
			}

			CloseClipboard();
		}
	}

	return result;
}

BITMAPINFO* CAddInClipboard::getImage()
{
	BITMAPINFO* result = NULL;

	if (IsClipboardFormatAvailable(CF_DIB))
	{
		if (OpenClipboard(NULL))
		{
			HANDLE hClipboard = GetClipboardData(CF_DIB);
			
			if (hClipboard != NULL && hClipboard != INVALID_HANDLE_VALUE)
			{
				void* dib = GlobalLock(hClipboard);

				if (dib)
				{
					result = (BITMAPINFO*)dib;

					GlobalUnlock(dib);
				}
			}

			CloseClipboard();
		}
	}

	return result;
}
