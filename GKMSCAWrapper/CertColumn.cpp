#include "StdAfx.h"
#include "CertColumn.h"

#include <Certsrv.h>

#pragma comment(lib, "ole32")	// COM
#pragma comment(lib, "Certidl")	// Certificate Services COM classes
#pragma comment(lib, "Certadm")

namespace GK {
namespace MSCAWrapper {

	CertColumn::CertColumn(System::String ^caConfiguration, System::String ^name)
	: MSCAObject(caConfiguration), _name(name)
	{
		_index = -1;
	}

	System::String ^CertColumn::name::get()
	{
		return _name;
	}

	void CertColumn::name::set(System::String ^value)
	{
		_name = value;
	}

	int CertColumn::index::get()
	{
		if (-1 == _index)
		{
			ICertView *pCertView = NULL;
			
			try
			{
				LONG lngIndex;
				pCertView = openCAConnection();

				BSTR bsColumnName = NULL;
				System::IntPtr pBstr = System::Runtime::InteropServices::Marshal::StringToBSTR(name);
				bsColumnName = reinterpret_cast<BSTR>(pBstr.ToPointer());

				pCertView->GetColumnIndex(CVRC_COLUMN_SCHEMA, bsColumnName, &lngIndex);
				
				System::Runtime::InteropServices::Marshal::FreeBSTR(pBstr);

				_index = lngIndex;
			}
			finally
			{
				CHECK_COM_NULL_AND_RELEASE(pCertView);
			}
		}

		return _index;
	}
}
}