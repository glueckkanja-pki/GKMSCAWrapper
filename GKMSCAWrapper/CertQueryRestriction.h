#pragma once

#include <Certsrv.h>

namespace GK {
namespace MSCAWrapper {
	public ref class CertQueryRestriction abstract
	{
		internal:
			/// <summary>
			/// The CertificateServices calls this to apply the filter to a given ICertView2 COM object
			/// </summary>
			virtual void restrictCertView(ICertView2 *view2Restrict) = 0;
	};
}
}