#pragma once
#include "certqueryrestriction.h"

namespace GK {
namespace MSCAWrapper {
	/// <summary>
	/// Applies multiple restrictions on a certificate query. Each restriction is represented by a CertQueryRestriction object
	/// </summary>
	public ref class IteratorRestriction : public CertQueryRestriction
	{
	public:
		property System::Collections::Generic::ICollection<CertQueryRestriction ^> ^QueryRestrictions;

		IteratorRestriction(System::Collections::Generic::IEnumerable<CertQueryRestriction ^> ^restrictions);

	internal:
		/// <summary>
		/// The CertificateServices calls this to apply the filter to a given ICertView2 COM object
		/// </summary>
		virtual void restrictCertView(ICertView2 *view2Restrict) override;
	};
}
}