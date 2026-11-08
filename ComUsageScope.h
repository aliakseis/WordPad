#pragma once

class CComUsageScope
{
    bool m_bInitialized;
public:
    explicit CComUsageScope(DWORD dwCoInit = COINIT_APARTMENTTHREADED)//COINIT_MULTITHREADED | COINIT_SPEED_OVER_MEMORY)
    {
        m_bInitialized = SUCCEEDED(CoInitializeEx(NULL, dwCoInit));
    }
    ~CComUsageScope()
    {
        if (m_bInitialized)
            CoUninitialize();
    }
};
