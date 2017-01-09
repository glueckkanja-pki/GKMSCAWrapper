#include "StdAfx.h"
#include "IteratorRestriction.h"

namespace GK {
namespace MSCAWrapper {
	IteratorRestriction::IteratorRestriction(System::Collections::Generic::IEnumerable<CertQueryRestriction ^> ^restrictions)
	{
		QueryRestrictions = gcnew System::Collections::Generic::LinkedList<CertQueryRestriction ^>(restrictions);
	}

	void IteratorRestriction::restrictCertView(ICertView2 *view2Restrict)
	{
		for each (CertQueryRestriction ^currentRestriction in QueryRestrictions)
			currentRestriction->restrictCertView(view2Restrict);
	}
}
}