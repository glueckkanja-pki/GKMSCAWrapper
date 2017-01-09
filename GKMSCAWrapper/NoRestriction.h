#pragma once
#include "CertQueryRestriction.h"

namespace GK {
	namespace MSCAWrapper {
		ref class NoRestriction :
			public CertQueryRestriction
		{
		public:
			NoRestriction();

			// Inherited via CertQueryRestriction
			virtual void restrictCertView(ICertView2 * view2Restrict) override;
		};
	}
}