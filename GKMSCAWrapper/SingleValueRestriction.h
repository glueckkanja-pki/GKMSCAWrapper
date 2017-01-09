#pragma once

#include "certqueryrestriction.h"
#include "CertColumn.h"

namespace GK {
namespace MSCAWrapper {
	/// <summary>
	/// Filters a query for certificates in the MS CA database for a single column that may either be
	/// required to have
	/// - values equivalent to a given value or
	/// - greater or smaller than a given value
	/// </summary>
	public ref class SingleValueRestriction : public CertQueryRestriction
	{
	public:
		enum class ComparisonOperator { Equals = CVR_SEEK_EQ, LowerThan = CVR_SEEK_LT, GreaterThan = CVR_SEEK_GT };

		/// <summary>
		/// This value is filtered for. Returned columns are required to be equal or in some other
		/// relationship to this value. The type of relationship is determined by the filterOperator
		/// </summary>
		property System::Object ^restrictionValue;

		/// <summary>
		/// Which column should be filtered for?
		/// </summary>
		property CertColumn ^restrictionColumn;

		/// <summary>
		/// In which kind of relation should the values be to restrictionValue? Default is equivalence (ComparisonOperator.Equals)
		/// </summary>
		property ComparisonOperator filterOperator;

	public:
		SingleValueRestriction(System::Object ^value2filter, CertColumn ^filterColumn);
	internal:

		/// <summary>
		/// The CertificateServices calls this to apply the filter to a given ICertView2 COM object
		/// </summary>
		virtual void restrictCertView(ICertView2 *view2Restrict) override;
	};
}
}