#include <windows.h>
#include <crtdbg.h>
#include <ActivScp.h>
#include <atlsafe.h>
#include <atlconv.h>
#include <map>

#include <lua.hpp>
#include <luacom.h>

#define OLESCRIPT_E_SYNTAX _HRESULT_TYPEDEF_(0x80020101L)

#ifdef _DEBUG
#   define TRACE( str, ... ) \
      { \
        char c[0xff]; \
        sprintf_s( c, str, __VA_ARGS__ ); \
        OutputDebugStringA( c ); \
      }
#else
#    define TRACE( str, ... )
#endif









const CLSID CLSID_LuaScript = {0x4B014691,0x33A9,0x429d,{0xA7,0x4D,0x4A,0x19,0xD9,0xAF,0xE8,0x4F}};	// 4B014691-33A9-429d-A74D-4A19D9AFE84F

long g_nRefCount = 0;









class CLuaScript
	: public IActiveScript
	, public IActiveScriptParse
	, public IActiveScriptError
	, public IDispatch
{
protected:
	long m_nRefCount;

	lua_State*			m_pLua;
	bool				m_bLuaLoaded;

	IActiveScriptSite*	m_pSite;
	std::map<CComBSTR, DWORD>	m_oGlobalNames;

	SCRIPTSTATE			m_eState;

	CComBSTR			m_sLastError;

public:
	CLuaScript()
		: m_nRefCount(1)
		, m_pSite(NULL)
		, m_eState(SCRIPTSTATE_UNINITIALIZED)
		, m_bLuaLoaded(false)
	{
		::InterlockedIncrement(&g_nRefCount);

		m_pLua = luaL_newstate();
		luaL_openlibs(m_pLua);
		luacom_open(m_pLua);
	}

	~CLuaScript()
	{
		luacom_close(m_pLua);
		lua_close(m_pLua);

		::InterlockedDecrement(&g_nRefCount);
	}

protected:
// IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface( 
        /* [in] */ REFIID riid,
        /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject)
	{
		if(::IsEqualGUID(IID_IUnknown, riid))
		{
			*ppvObject = static_cast<IUnknown*>(static_cast<IDispatch*>(this));
		}
		else
		if(::IsEqualGUID(__uuidof(IDispatch), riid))
		{
			*ppvObject = static_cast<IDispatch*>(this);
		}
		else
		if(::IsEqualGUID(__uuidof(IActiveScript), riid))
		{
			*ppvObject = static_cast<IActiveScript*>(this);
		}
		else
		if(::IsEqualGUID(__uuidof(IActiveScriptParse), riid))
		{
			*ppvObject = static_cast<IActiveScriptParse*>(this);
		}
		else
		{
			return E_NOINTERFACE;
		}

		AddRef();

		return S_OK;
	}
    
    virtual ULONG STDMETHODCALLTYPE AddRef( void)
	{
		return ++m_nRefCount;
	}
    
    virtual ULONG STDMETHODCALLTYPE Release( void)
	{
		if(!--m_nRefCount)
		{
			delete this;
			return 0;
		}

		return m_nRefCount;
	}

// IDispatch
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount( 
        /* [out] */ UINT *pctinfo)
	{
		return E_NOTIMPL;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetTypeInfo( 
        /* [in] */ UINT iTInfo,
        /* [in] */ LCID lcid,
        /* [out] */ ITypeInfo **ppTInfo)
	{
		return E_NOTIMPL;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames( 
        /* [in] */ REFIID riid,
        /* [size_is][in] */ LPOLESTR *rgszNames,
        /* [in] */ UINT cNames,
        /* [in] */ LCID lcid,
        /* [size_is][out] */ DISPID *rgDispId)
	{
		return E_NOTIMPL;
	}
    
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE Invoke( 
        /* [in] */ DISPID dispIdMember,
        /* [in] */ REFIID riid,
        /* [in] */ LCID lcid,
        /* [in] */ WORD wFlags,
        /* [out][in] */ DISPPARAMS *pDispParams,
        /* [out] */ VARIANT *pVarResult,
        /* [out] */ EXCEPINFO *pExcepInfo,
        /* [out] */ UINT *puArgErr)
	{
		return E_NOTIMPL;
	}

// IActiveScript
    virtual HRESULT STDMETHODCALLTYPE SetScriptSite( 
        /* [in] */ __RPC__in_opt IActiveScriptSite *pass)
	{
		if(m_pSite)
			m_pSite->Release();

		(m_pSite = pass)->AddRef();

		return S_OK;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetScriptSite( 
        /* [in] */ __RPC__in REFIID riid,
        /* [iid_is][out] */ __RPC__deref_out_opt void **ppvObject)
	{
		return E_NOTIMPL;
	}
    
    virtual HRESULT STDMETHODCALLTYPE SetScriptState( 
        /* [in] */ SCRIPTSTATE ss)
	{
		HRESULT hr = S_OK;

		if(ss == SCRIPTSTATE_STARTED){
			hr = m_bLuaLoaded ? Run() : S_OK;
		}
		else
		if(ss == SCRIPTSTATE_UNINITIALIZED){
			for(std::map<CComBSTR,DWORD>::iterator i=m_oGlobalNames.begin(); i!=m_oGlobalNames.end(); ++i){
				lua_pushnil(m_pLua);
				lua_setglobal(m_pLua, ATL::CW2A(i->first));
			}

			if(m_pSite){
				m_pSite->Release();
				m_pSite = NULL;
			}
		}

		m_eState = ss;

		return hr;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetScriptState( 
        /* [out] */ __RPC__out SCRIPTSTATE *pssState)
	{
		*pssState = m_eState;

		return S_OK;
	}
    
    virtual HRESULT STDMETHODCALLTYPE Close( void)
	{
		return E_NOTIMPL;
	}
    
    virtual HRESULT STDMETHODCALLTYPE AddNamedItem( 
        /* [in] */ __RPC__in LPCOLESTR pstrName,
        /* [in] */ DWORD dwFlags)
	{
		m_oGlobalNames.insert(std::map<CComBSTR,DWORD>::value_type(pstrName, dwFlags));

		return S_OK;
	}
    
    virtual HRESULT STDMETHODCALLTYPE AddTypeLib( 
        /* [in] */ __RPC__in REFGUID rguidTypeLib,
        /* [in] */ DWORD dwMajor,
        /* [in] */ DWORD dwMinor,
        /* [in] */ DWORD dwFlags)
	{
		return E_NOTIMPL;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetScriptDispatch( 
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [out] */ __RPC__deref_out_opt IDispatch **ppdisp)
	{
		(*ppdisp = static_cast<IDispatch*>(this))->AddRef();

		return S_OK;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetCurrentScriptThreadID( 
        /* [out] */ __RPC__out SCRIPTTHREADID *pstidThread)
	{
		return E_NOTIMPL;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetScriptThreadID( 
        /* [in] */ DWORD dwWin32ThreadId,
        /* [out] */ __RPC__out SCRIPTTHREADID *pstidThread)
	{
		return E_NOTIMPL;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetScriptThreadState( 
        /* [in] */ SCRIPTTHREADID stidThread,
        /* [out] */ __RPC__out SCRIPTTHREADSTATE *pstsState)
	{
		return E_NOTIMPL;
	}
    
    virtual HRESULT STDMETHODCALLTYPE InterruptScriptThread( 
        /* [in] */ SCRIPTTHREADID stidThread,
        /* [in] */ __RPC__in const EXCEPINFO *pexcepinfo,
        /* [in] */ DWORD dwFlags)
	{
		return E_NOTIMPL;
	}
    
    virtual HRESULT STDMETHODCALLTYPE Clone( 
        /* [out] */ __RPC__deref_out_opt IActiveScript **ppscript)
	{
		return E_NOTIMPL;
	}

// IActiveScriptParse
    virtual HRESULT STDMETHODCALLTYPE InitNew( void)
	{
		return S_OK;
	}
    
	// x86
    virtual HRESULT STDMETHODCALLTYPE AddScriptlet( 
        /* [in] */ __RPC__in LPCOLESTR pstrDefaultName,
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in LPCOLESTR pstrSubItemName,
        /* [in] */ __RPC__in LPCOLESTR pstrEventName,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORD dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__deref_out_opt BSTR *pbstrName,
        /* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
	{
		return AddScriptlet(
			pstrDefaultName, 
			pstrCode, 
			pstrItemName, 
			pstrSubItemName, 
			pstrEventName, 
			pstrDelimiter, 
			(DWORDLONG)dwSourceContextCookie, 
			ulStartingLineNumber, 
			dwFlags, 
			pbstrName, 
			pexcepinfo
		);
	}
    
	// x86
    virtual HRESULT STDMETHODCALLTYPE ParseScriptText( 
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in_opt IUnknown *punkContext,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORD dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__out VARIANT *pvarResult,
        /* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
	{
		return ParseScriptText( 
			pstrCode,
			pstrItemName,
			punkContext,
			pstrDelimiter,
			(DWORDLONG)dwSourceContextCookie,
			ulStartingLineNumber,
			dwFlags,
			pvarResult,
			pexcepinfo
		);
	}

	// x64
    virtual HRESULT STDMETHODCALLTYPE AddScriptlet( 
        /* [in] */ __RPC__in LPCOLESTR pstrDefaultName,
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in LPCOLESTR pstrSubItemName,
        /* [in] */ __RPC__in LPCOLESTR pstrEventName,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORDLONG dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__deref_out_opt BSTR *pbstrName,
        /* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
	{
		return E_NOTIMPL;
	}
    
	// x64
    virtual HRESULT STDMETHODCALLTYPE ParseScriptText( 
        /* [in] */ __RPC__in LPCOLESTR pstrCode,
        /* [in] */ __RPC__in LPCOLESTR pstrItemName,
        /* [in] */ __RPC__in_opt IUnknown *punkContext,
        /* [in] */ __RPC__in LPCOLESTR pstrDelimiter,
        /* [in] */ DWORDLONG dwSourceContextCookie,
        /* [in] */ ULONG ulStartingLineNumber,
        /* [in] */ DWORD dwFlags,
        /* [out] */ __RPC__out VARIANT *pvarResult,
        /* [out] */ __RPC__out EXCEPINFO *pexcepinfo)
	{
		HRESULT hr = S_OK;

		size_t nRequired = ::WideCharToMultiByte(CP_UTF8, 0, pstrCode, -1, NULL, 0, NULL, NULL);

		char* pCode = new char[nRequired+1];
		pCode[nRequired] = '\0';

		WideCharToMultiByte(CP_UTF8, 0, pstrCode, -1, pCode, (int)nRequired, NULL, NULL);

		if(luaL_loadstring(m_pLua, pCode)){
			m_sLastError = lua_tostring(m_pLua, -1);
			m_pSite->OnScriptError( static_cast<IActiveScriptError*>(this) );
		}else{
//			lua_setglobal(m_pLua, "_");
			m_bLuaLoaded = true;
		}

		delete[] pCode;

		if(m_bLuaLoaded && m_eState == SCRIPTSTATE_STARTED){
			hr = Run();
		}

		return hr;
	}

// IActiveScriptError
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetExceptionInfo( 
        /* [out] */ EXCEPINFO *pexcepinfo)
	{
		::memset(pexcepinfo, 0, sizeof(EXCEPINFO));
		pexcepinfo->bstrSource = ::SysAllocString(L"LuaScript");
		pexcepinfo->bstrDescription = ::SysAllocString(m_sLastError);
		pexcepinfo->scode = OLESCRIPT_E_SYNTAX;

		return S_OK;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetSourcePosition( 
        /* [out] */ __RPC__out DWORD *pdwSourceContext,
        /* [out] */ __RPC__out ULONG *pulLineNumber,
        /* [out] */ __RPC__out LONG *plCharacterPosition)
	{
		if(pdwSourceContext) *pdwSourceContext = 0;
		*pulLineNumber = 0;
		*plCharacterPosition = 0;
		return S_OK;
	}
    
    virtual HRESULT STDMETHODCALLTYPE GetSourceLineText( 
        /* [out] */ __RPC__deref_out_opt BSTR *pbstrSourceLine)
	{
		*pbstrSourceLine = NULL;
		return E_NOTIMPL;
	}

// CLuaScript
	HRESULT Run(){
		for(std::map<CComBSTR,DWORD>::iterator i=m_oGlobalNames.begin(); i!=m_oGlobalNames.end(); ++i){
			LPCOLESTR pstrName = (i->first).m_str;

			IUnknown* pUnknown = NULL;
			ITypeInfo* pTypeInfo = NULL;
			IDispatch* pDispatch = NULL;

			if( 
				SUCCEEDED(m_pSite->GetItemInfo(pstrName, SCRIPTINFO_IUNKNOWN, &pUnknown, &pTypeInfo))	&&
				SUCCEEDED(pUnknown->QueryInterface(__uuidof(IDispatch), (void**)&pDispatch))			&&
				true)
			{
				luacom_IDispatch2LuaCOM(m_pLua, pDispatch);
				lua_setglobal(m_pLua, ATL::CW2A(pstrName));
			}

			if(pUnknown) pUnknown->Release();
			if(pTypeInfo) pTypeInfo->Release();
			if(pDispatch) pDispatch->Release();
		}

		{
			m_pSite->OnEnterScript();

			if(lua_pushvalue(m_pLua, 1), lua_pcall(m_pLua, 0, 0, 0)){
				m_sLastError = lua_tostring(m_pLua, -1);
				m_pSite->OnScriptError( static_cast<IActiveScriptError*>(this) );
			}

			m_pSite->OnLeaveScript();
		}

		return S_OK;
	}

};
//---



class CClassFactoryLuaScript
	: public IClassFactory
{
protected:
// IUnknown
	virtual HRESULT STDMETHODCALLTYPE QueryInterface( 
        /* [in] */ REFIID riid,
        /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject)
	{
		if(::IsEqualGUID(IID_IUnknown, riid))
		{
			*ppvObject = static_cast<IUnknown*>(this);
		}
		else
		if(::IsEqualGUID(IID_IClassFactory, riid))
		{
			*ppvObject = static_cast<IClassFactory*>(this);
		}
		else
		{
			return E_NOINTERFACE;
		}

		AddRef();

		return S_OK;
	}
    
    virtual ULONG STDMETHODCALLTYPE AddRef( void)
	{
		return ::InterlockedIncrement(&g_nRefCount);
	}
    
    virtual ULONG STDMETHODCALLTYPE Release( void)
	{
		return ::InterlockedDecrement(&g_nRefCount);
	}
    
// IClassFactory
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE CreateInstance( 
        /* [unique][in] */ IUnknown *pUnkOuter,
        /* [in] */ REFIID riid,
        /* [iid_is][out] */ void **ppvObject)
	{
		if(pUnkOuter) return CLASS_E_NOAGGREGATION;

		IUnknown* pTarget = static_cast<IUnknown*>(static_cast<IActiveScript*>(new CLuaScript));
		if(!pTarget) return E_OUTOFMEMORY;

		HRESULT hr = pTarget->QueryInterface(riid, ppvObject);
		
		pTarget->Release();

		return hr;
	}
    
    virtual /* [local] */ HRESULT STDMETHODCALLTYPE LockServer( 
        /* [in] */ BOOL fLock)
	{
		return S_OK;
	}

};
//---









HINSTANCE	g_hDLL						= NULL;
wchar_t		g_aDLLPath[MAX_PATH]		= L"";



BOOL WINAPI DllMain(
	HINSTANCE hinstDLL,
	DWORD fdwReason,
	LPVOID lpvReserved)
{
	BOOL bReturn = TRUE;

	switch(fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		{
			g_hDLL = hinstDLL;
			::GetModuleFileNameW(g_hDLL, g_aDLLPath, MAX_PATH);



			::DisableThreadLibraryCalls(hinstDLL);



		    #if defined(_DEBUG)
	        // メモリリーク検出用
	        ::_CrtSetDbgFlag(
	              _CRTDBG_ALLOC_MEM_DF
	            | _CRTDBG_LEAK_CHECK_DF
			 );
			new int(0x12345678);
		    #endif
		}
		break;
	case DLL_PROCESS_DETACH:
		{
			// none
		}
		break;
	}

	return bReturn;
}
//---









CClassFactoryLuaScript g_oCF_LuaScript;



HRESULT RegistCOM(const wchar_t* pProgID, const wchar_t* pCLSID, const wchar_t* pThreadingModel)
{
	HRESULT hr = E_FAIL;



	HKEY hPROGID		= NULL;
	HKEY hPROGID_CLSID	= NULL;
	HKEY hCLSID			= NULL;
	HKEY hCLSID_CLSID	= NULL;
	HKEY hCLSID_CLSID_InProcServer32	= NULL;

	DWORD dwDisposition;

	if(
		::RegCreateKeyExW(HKEY_CLASSES_ROOT, pProgID, 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hPROGID, &dwDisposition)	== ERROR_SUCCESS &&
		::RegCreateKeyExW(hPROGID, L"CLSID", 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hPROGID_CLSID, &dwDisposition)		== ERROR_SUCCESS &&
		::RegSetValueExW(hPROGID_CLSID, NULL, 0, REG_SZ, (const BYTE *)pCLSID, 38*2)														== ERROR_SUCCESS &&

		::RegOpenKeyExW(HKEY_CLASSES_ROOT, L"CLSID", 0, KEY_READ | KEY_WRITE, &hCLSID)														== ERROR_SUCCESS &&
			::RegCreateKeyExW(hCLSID, pCLSID, 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hCLSID_CLSID, &dwDisposition)									== ERROR_SUCCESS &&
			::RegSetValueExW(hCLSID_CLSID, NULL, 0, REG_SZ, (const BYTE *)pProgID, (DWORD)::wcslen(pProgID)*2)															== ERROR_SUCCESS &&
			::RegCreateKeyExW(hCLSID_CLSID, L"InProcServer32", 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hCLSID_CLSID_InProcServer32, &dwDisposition)	== ERROR_SUCCESS &&
			::RegSetValueExW(hCLSID_CLSID_InProcServer32, NULL, 0, REG_SZ, (const BYTE *)g_aDLLPath, (DWORD)::wcslen(g_aDLLPath)*2)										== ERROR_SUCCESS &&
			::RegSetValueExW(hCLSID_CLSID_InProcServer32, L"ThreadingModel", 0, REG_SZ, (const BYTE *)pThreadingModel, (DWORD)::wcslen(pThreadingModel)*2)				== ERROR_SUCCESS &&
		true)
	{
		hr = S_OK;
	}

	if(hPROGID)						::RegCloseKey(hPROGID);
	if(hPROGID_CLSID)				::RegCloseKey(hPROGID_CLSID);
	if(hCLSID)						::RegCloseKey(hCLSID);
	if(hCLSID_CLSID)				::RegCloseKey(hCLSID_CLSID);
	if(hCLSID_CLSID_InProcServer32)	::RegCloseKey(hCLSID_CLSID_InProcServer32);



	return hr;
}

HRESULT UnregistCOM(const wchar_t* pProgID, const wchar_t* pCLSID)
{
	HRESULT hr = E_FAIL;



	HKEY hPROGID		= NULL;
	HKEY hCLSID			= NULL;
	HKEY hCLSID_CLSID	= NULL;

	if(
		::RegOpenKeyExW(HKEY_CLASSES_ROOT, pProgID, 0, KEY_READ | KEY_WRITE, &hPROGID)	== ERROR_SUCCESS &&
		::RegDeleteKeyW(hPROGID, L"CLSID")												== ERROR_SUCCESS &&
		::RegDeleteKeyW(HKEY_CLASSES_ROOT, pProgID)										== ERROR_SUCCESS &&

		::RegOpenKeyExW(HKEY_CLASSES_ROOT, L"CLSID", 0, KEY_READ | KEY_WRITE, &hCLSID)	== ERROR_SUCCESS &&
			::RegOpenKeyExW(hCLSID, pCLSID, 0, KEY_READ | KEY_WRITE, &hCLSID_CLSID)			== ERROR_SUCCESS &&
			::RegDeleteKeyW(hCLSID_CLSID, L"InProcServer32")								== ERROR_SUCCESS &&
			::RegDeleteKeyW(hCLSID, pCLSID)													== ERROR_SUCCESS &&
		true)
	{
		hr = S_OK;
	}

	if(hPROGID)			::RegCloseKey(hPROGID);
	if(hCLSID)			::RegCloseKey(hCLSID);
	if(hCLSID_CLSID)	::RegCloseKey(hCLSID_CLSID);



	return hr;
}



STDAPI  DllGetClassObject(IN REFCLSID rclsid, IN REFIID riid, OUT LPVOID FAR* ppv)
{
	if(::IsEqualGUID(rclsid, CLSID_NULL))
	{
		wchar_t* pClassName = (wchar_t*)*ppv;

		if(!pClassName || pClassName[0] == L'\0')
		{
			return ((IClassFactory*)&g_oCF_LuaScript)->QueryInterface(riid, ppv);
		}
		else
		if(::wcscmp(pClassName, L"LuaScript") == 0)
		{
			return ((IClassFactory*)&g_oCF_LuaScript)->QueryInterface(riid, ppv);
		}
		else
		{
			return CLASS_E_CLASSNOTAVAILABLE;
		}
	}
	else
	if(::IsEqualGUID(rclsid, CLSID_LuaScript))
	{
		return ((IClassFactory*)&g_oCF_LuaScript)->QueryInterface(riid, ppv);
	}
	else
	{
		return CLASS_E_CLASSNOTAVAILABLE;
	}

}

STDAPI  DllCanUnloadNow(void)
{
	return g_nRefCount ? S_FALSE : S_OK;
}

STDAPI DllRegisterServer(void)
{
	HRESULT hr = E_FAIL;



	HKEY hW3SVC			= NULL;
	HKEY hW3SVC_ASP		= NULL;
	HKEY hW3SVC_ASP_LanguageEngines			= NULL;
	HKEY hW3SVC_ASP_LanguageEngines_Script	= NULL;
	HKEY hLua					= NULL;
	HKEY hLuaFile				= NULL;
	HKEY hLuaFile_ScriptEngine	= NULL;
	HKEY hLuaFile_DefaultIcon	= NULL;

	DWORD dwDisposition;

	if(
		SUCCEEDED(hr = RegistCOM(L"LuaScript",						L"{4B014691-33A9-429d-A74D-4A19D9AFE84F}", L"Apartment"))	&&

		::RegCreateKeyExW(HKEY_CLASSES_ROOT, L".lua", 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hLua, &dwDisposition)			== ERROR_SUCCESS &&
			::RegSetValueExW(hLua, NULL, 0, REG_SZ, (const BYTE *)L"LuaFile", (DWORD)::wcslen(L"LuaFile")*2)																		== ERROR_SUCCESS &&
		::RegCreateKeyExW(HKEY_CLASSES_ROOT, L"LuaFile", 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hLuaFile, &dwDisposition)	== ERROR_SUCCESS &&
			::RegCreateKeyExW(hLuaFile, L"ScriptEngine", 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hLuaFile_ScriptEngine, &dwDisposition)							== ERROR_SUCCESS &&
			::RegSetValueExW(hLuaFile_ScriptEngine, NULL, 0, REG_SZ, (const BYTE *)L"LuaScript", (DWORD)::wcslen(L"LuaScript")*2)													== ERROR_SUCCESS &&
			::RegCreateKeyExW(hLuaFile, L"DefaultIcon", 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hLuaFile_DefaultIcon, &dwDisposition)							== ERROR_SUCCESS &&
			::RegSetValueExW(hLuaFile_DefaultIcon, NULL, 0, REG_SZ, (const BYTE *)L"C:\\Windows\\System32\\WScript.exe,1", (DWORD)::wcslen(L"C:\\Windows\\System32\\WScript.exe,1")*2)	== ERROR_SUCCESS &&

		::RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SYSTEM\\CurrentControlSet\\services\\W3SVC\\", 0, KEY_READ | KEY_WRITE, &hW3SVC)					== ERROR_SUCCESS &&
			::RegCreateKeyExW(hW3SVC, L"ASP", 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hW3SVC_ASP, &dwDisposition)												== ERROR_SUCCESS &&
			::RegCreateKeyExW(hW3SVC_ASP, L"LanguageEngines", 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hW3SVC_ASP_LanguageEngines, &dwDisposition)				== ERROR_SUCCESS &&
			::RegCreateKeyExW(hW3SVC_ASP_LanguageEngines, L"LuaScript", 0, REG_NONE, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, NULL, &hW3SVC_ASP_LanguageEngines_Script, &dwDisposition)	== ERROR_SUCCESS &&
			::RegSetValueExW(hW3SVC_ASP_LanguageEngines_Script, L"Write", 0, REG_SZ, (const BYTE *)L"Response:Write(|)", (DWORD)::wcslen(L"Response:Write(|)")*2)					== ERROR_SUCCESS &&
			::RegSetValueExW(hW3SVC_ASP_LanguageEngines_Script, L"WriteBlock", 0, REG_SZ, (const BYTE *)L"Response:WriteBlock(|)", (DWORD)::wcslen(L"Response:WriteBlock(|)")*2)	== ERROR_SUCCESS &&
		true)
	{
		hr = S_OK;
	}

	if(hW3SVC)						::RegCloseKey(hW3SVC);
	if(hW3SVC_ASP)					::RegCloseKey(hW3SVC_ASP);
	if(hW3SVC_ASP_LanguageEngines)	::RegCloseKey(hW3SVC_ASP_LanguageEngines);
	if(hW3SVC_ASP_LanguageEngines_Script)	::RegCloseKey(hW3SVC_ASP_LanguageEngines_Script);
	if(hLua)						::RegCloseKey(hLua);
	if(hLuaFile)					::RegCloseKey(hLuaFile);
	if(hLuaFile_ScriptEngine)		::RegCloseKey(hLuaFile_ScriptEngine);
	if(hLuaFile_DefaultIcon)		::RegCloseKey(hLuaFile_DefaultIcon);



	return hr;
}

STDAPI DllUnregisterServer(void)
{
	HRESULT hr = E_FAIL;



	if(
		SUCCEEDED(hr = UnregistCOM(L"LuaScript",					L"{4B014691-33A9-429d-A74D-4A19D9AFE84F}"))	&&
		true)
	{
		hr = S_OK;
	}



	return hr;
}








