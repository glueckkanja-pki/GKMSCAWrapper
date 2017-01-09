#include "StdAfx.h"
#include "MSCAObject.h"
#include <Certsrv.h>

#pragma comment(lib, "ole32")	// COM
#pragma comment(lib, "Certidl")	// Certificate Services COM classes
#pragma comment(lib, "Certadm")

namespace GK {
namespace MSCAWrapper {

	MSCAObject::MSCAObject(System::String ^caConfiguration)
	{
		this->caConfiguration = caConfiguration;

		HRESULT hr = CoInitializeEx(NULL, NULL);
		ASSERT_COM_SUCCESS(hr, "Couldn't initialize COM subsystem");
	}

		// TODO: CoUninitialize

	ICertView2 *MSCAObject::openCAConnection()
	{
		ICertView2 *pCertView = NULL;
		HRESULT hr;

		try
		{
			hr = CoCreateInstance(CLSID_CCertView,
						  NULL,
						  CLSCTX_INPROC_SERVER,
						  IID_ICertView2,
						  (void **)&pCertView);
			ASSERT_COM_SUCCESS(hr, "Could not create COM object for CA connection");

				// Initialize Connection on pCertView
			BSTR bsCaConfiguration = NULL;
			System::IntPtr pBstr = System::Runtime::InteropServices::Marshal::StringToBSTR(caConfiguration);
			bsCaConfiguration = reinterpret_cast<BSTR>(pBstr.ToPointer());
			hr = pCertView->OpenConnection(bsCaConfiguration);
			System::Runtime::InteropServices::Marshal::FreeBSTR(pBstr);
			bsCaConfiguration = NULL;
			if (0x800706BA == hr)
				throw gcnew System::InvalidOperationException("CA server could not be contacted via RPC (in fact, it probably doesn't even exist)");
			if (0x80040154 ==  hr)
				throw gcnew System::InvalidOperationException("Server was found, but Certificate Services doesn't seem to be installed.");
			if (0x80070057 == hr)
				throw gcnew System::ArgumentException(
				System::String::Concat(	
					"The server does not contain Certificate Services with the requested name \"",
					caConfiguration,
					"\""));
			if (0x80070005 == hr)
				throw gcnew System::UnauthorizedAccessException("Access to Certification Authority was denied. No permission to connect to the CA.");
			if (0x80070103 == hr)
				throw gcnew System::InvalidOperationException("The machine you are trying to connect to the certification authority from is not in the same domain as the certificate services system");
			// Also seen: 0x80070006 on situations where the CA is in another domain than the client the tool is running on.


			ASSERT_COM_SUCCESS(hr, "Could not open connection to CA");

			return pCertView;
		}
		catch(...)
		{
			CHECK_COM_NULL_AND_RELEASE(pCertView);
			throw;
		}
	}
}
}