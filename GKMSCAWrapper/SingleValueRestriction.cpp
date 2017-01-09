#include "StdAfx.h"
#include "SingleValueRestriction.h"

namespace GK {
namespace MSCAWrapper {
	SingleValueRestriction::SingleValueRestriction(System::Object ^value2filter, CertColumn ^filterColumn)
	{
		restrictionColumn = filterColumn;
		restrictionValue = value2filter;
		filterOperator = ComparisonOperator::Equals;
	}

	void SingleValueRestriction::restrictCertView(ICertView2 *view2Restrict)
	{
		VARIANT varRestrictionValue;
		System::Runtime::InteropServices::Marshal::GetNativeVariantForObject(restrictionValue, System::IntPtr(&varRestrictionValue));
		HRESULT hr = view2Restrict->SetRestriction(restrictionColumn->index, (LONG)filterOperator, CVR_SORT_NONE, &varRestrictionValue);
		VariantClear(&varRestrictionValue);
		ASSERT_COM_SUCCESS(hr, "Could not restrict certificate view to issued (or other specific) certificates");
	}
}
}