// dllmain.h : Declaration of module class.

class CSimpleExtModule : public ATL::CAtlDllModuleT< CSimpleExtModule >
{
public :
	DECLARE_LIBID(LIBID_SimpleExtLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_SIMPLEEXT, "{1048A6A2-F558-4C61-B149-10C060DC5C28}")
};

extern class CSimpleExtModule _AtlModule;
