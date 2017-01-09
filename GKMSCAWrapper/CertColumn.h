#pragma once
#include "mscaobject.h"

namespace GK {
namespace MSCAWrapper {

	public ref class CertColumn : public MSCAObject
	{
	public:
		property System::String ^name{
			System::String ^get();
			private: void set(System::String ^value);
		}

	private:
		int _index;
	internal:
		property int index{
			int get();
		}

		System::String ^_name;

		CertColumn(System::String ^caConfiguration, System::String ^name);
	};

}
}