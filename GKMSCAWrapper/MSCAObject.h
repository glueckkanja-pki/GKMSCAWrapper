#pragma once

struct ICertView2;

namespace GK {
namespace MSCAWrapper {

	public ref class MSCAObject
	{
	public:
		MSCAObject(System::String ^caConfiguration);
	protected:
		System::String ^caConfiguration;

		ICertView2 *openCAConnection();
	};

}}