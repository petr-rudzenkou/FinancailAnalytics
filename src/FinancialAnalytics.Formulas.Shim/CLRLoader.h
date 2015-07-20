// CLRLoader.h
#pragma once

class CCLRLoader
{
public:
    CCLRLoader(void);
    virtual ~CCLRLoader(void);

    HRESULT CreateAggregatedAddIn(
        IUnknown* pOuter,
        LPCWSTR szAssemblyName, 
        LPCWSTR szClassName, 
        LPCWSTR szAssemblyConfigName);
    HRESULT Unload(void);

	//void GetFormulasObject(FinancialAnalytics_Formulas::IFormulasProcessorPtr &formulasObject);

private:
    HRESULT LoadCLR();
    HRESULT CreateAppDomain(LPCWSTR szAssemblyConfigName);

    ICorRuntimeHost *m_pCorRuntimeHost;
    mscorlib::_AppDomain *m_pAppDomain;
};
