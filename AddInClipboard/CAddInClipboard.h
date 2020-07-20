#pragma once

#include "ComponentBase.h"
#include "AddInDefBase.h"
#include "IMemoryManager.h"

class CAddInClipboard : public IComponentBase
{
public:
	CAddInClipboard(void);
	~CAddInClipboard();
	
	// IInitDoneBase
	bool ADDIN_API Init(void* disp);
	bool ADDIN_API setMemManager(void* memManager);
	long ADDIN_API GetInfo();
	void ADDIN_API Done();

	// ILanguageExtenderBase
	bool ADDIN_API RegisterExtensionAs(WCHAR_T** wsExtName);
	long ADDIN_API GetNProps();
	long ADDIN_API FindProp(const WCHAR_T* wsPropName);
	const WCHAR_T* ADDIN_API GetPropName(long lPropNum, long lPropAlias);
	bool ADDIN_API GetPropVal(const long lPropNum, tVariant* pvarPropVal);
	bool ADDIN_API SetPropVal(const long lPropNum, tVariant* pvarPropVal);
	bool ADDIN_API IsPropReadable(const long lPropNum);
	bool ADDIN_API IsPropWritable(const long lPropNum);
	long ADDIN_API GetNMethods();
	long ADDIN_API FindMethod(const WCHAR_T* wsMethodName);
	const WCHAR_T* ADDIN_API GetMethodName(const long lMethodNum, const long lMethodAlias);
	long ADDIN_API GetNParams(const long lMethodNum);
	bool ADDIN_API GetParamDefValue(const long lMethodNum, const long lParamNum, tVariant* pvarParamDefValue);
	bool ADDIN_API HasRetVal(const long lMethodNum);
	bool ADDIN_API CallAsProc(const long lMethodNum, tVariant* paParams, const long lSizeArray);
	bool ADDIN_API CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray);
	
	// LocaleBase
	void ADDIN_API SetLocale(const WCHAR_T* wsLocale);

private:
	// Attributes
	IAddInDefBase* m_iConnect;
	IMemoryManager* m_iMemory;

	bool containsText();
	bool containsImage();
	const wchar_t* getText();
	BITMAPINFO* getImage();
};